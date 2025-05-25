#define COBJMACROS
#include "wincontrol.h"
#include <initguid.h>
#include <UIAutomation.h>
#include <stdio.h>
#include <tlhelp32.h>
#include <psapi.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <stdbool.h>

#define VK_DELETE 0x2E
#define VK_HOME   0x24
#define VK_END    0x23
#define VK_PAGEUP 0x21
#define VK_PAGEDOWN 0x22
#define VK_F1     0x70
#define VK_F2     0x71
#define VK_F3     0x72
#define VK_F4     0x73
#define VK_F5     0x74
#define VK_F6     0x75
#define VK_F7     0x76
#define VK_F8     0x77
#define VK_F9     0x78
#define VK_F10    0x79
#define VK_F11    0x7A
#define VK_F12    0x7B

struct EnumData {
    DWORD process_id;
    HWND window;
};

bool evaluate_condition(WinControlContext* ctx, const char* condition);
typedef bool (*CommandHandler)(WinControlContext* ctx, const Command* cmd);

bool winctrl_initialize(WinControlContext* ctx) {
    printf("Initializing COM...\n");

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Failed to initialize COM: 0x%lx", hr);
        printf("COM initialization failed: 0x%lx\n", hr);
        return false;
    }
    printf("COM initialized successfully\n");

    printf("Creating UI Automation instance...\n");
    hr = CoCreateInstance(&CLSID_CUIAutomation, NULL,
        CLSCTX_INPROC_SERVER, &IID_IUIAutomation,
        (void**)&ctx->automation);

    if (FAILED(hr)) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Failed to create UI Automation instance: 0x%lx (Try running as administrator)", hr);
        printf("UI Automation creation failed: 0x%lx\n", hr);
        CoUninitialize();
        return false;
    }

    printf("UI Automation instance created successfully\n");

    ctx->current_process_id = 0;
    ctx->current_window = NULL;
    ctx->last_error[0] = '\0';
    ctx->vars.variable_count = 0;
    ctx->typing_delay_ms = 0;
    ctx->log_file = NULL;
    ctx->log_filename[0] = '\0';


    return true;
}

static void print_window_info(HWND hwnd) {
    char title[256] = {0};
    char class_name[256] = {0};
    GetWindowTextA(hwnd, title, sizeof(title));
    GetClassNameA(hwnd, class_name, sizeof(class_name));
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    printf("Window found - Title: '%s', Class: '%s', PID: %lu\n",
           title, class_name, pid);
}

static BOOL CALLBACK enum_windows_callback(HWND hwnd, LPARAM lParam) {
    struct EnumData* data = (struct EnumData*)lParam;
    DWORD window_process_id;
    GetWindowThreadProcessId(hwnd, &window_process_id);

    if (window_process_id == data->process_id) {
        print_window_info(hwnd);
        if (IsWindowVisible(hwnd) && !IsIconic(hwnd)) {
            char title[256] = {0};
            GetWindowTextA(hwnd, title, sizeof(title));

            if (strlen(title) > 0) {
                LONG_PTR style = GetWindowLongPtr(hwnd, GWL_STYLE);
                if (style & WS_OVERLAPPEDWINDOW) {
                    data->window = hwnd;
                    return FALSE;
                }
            }
        }
    }
    return TRUE;
}


void winctrl_cleanup(WinControlContext* ctx) {
    if (ctx->automation) {
        IUIAutomation_Release(ctx->automation);
        ctx->automation = NULL;
    }
    CoUninitialize();
}

static IUIAutomationCondition* create_name_condition(WinControlContext* ctx, const char* name) {
    IUIAutomationCondition* condition = NULL;
    WCHAR wide_name[256];
    MultiByteToWideChar(CP_UTF8, 0, name, -1, wide_name, 256);

    VARIANT var;
    var.vt = VT_BSTR;
    var.bstrVal = SysAllocString(wide_name);

    ctx->automation->lpVtbl->CreatePropertyCondition(
        ctx->automation,
        UIA_NamePropertyId,
        var,
        &condition
    );

    SysFreeString(var.bstrVal);
    return condition;
}



static IUIAutomationElement* get_root_element(WinControlContext* ctx) {
    IUIAutomationElement* root = NULL;
    ctx->automation->lpVtbl->ElementFromHandle(
        ctx->automation,
        ctx->current_window,
        &root
    );
    return root;
}

bool winctrl_set_variable(WinControlContext* ctx, const char* name, const char* value) {
    for (int i = 0; i < ctx->vars.variable_count; i++) {
        if (strcmp(ctx->vars.variables[i].name, name) == 0) {
            strncpy_s(ctx->vars.variables[i].value,
                     sizeof(ctx->vars.variables[i].value),
                     value, _TRUNCATE);
            return true;
        }
    }

    if (ctx->vars.variable_count < MAX_VARIABLES) {
        strncpy_s(ctx->vars.variables[ctx->vars.variable_count].name,
                 sizeof(ctx->vars.variables[ctx->vars.variable_count].name),
                 name, _TRUNCATE);
        strncpy_s(ctx->vars.variables[ctx->vars.variable_count].value,
                 sizeof(ctx->vars.variables[ctx->vars.variable_count].value),
                 value, _TRUNCATE);
        ctx->vars.variable_count++;
        return true;
    }

    return false;
}

bool winctrl_get_element_text(IUIAutomationElement* element, char* text, size_t text_size) {
    if (!element || !text) return false;

    BSTR name = NULL;
    HRESULT hr = element->lpVtbl->get_CurrentName(element, &name);

    if (SUCCEEDED(hr) && name) {
        WideCharToMultiByte(CP_UTF8, 0, name, -1, text, (int)text_size, NULL, NULL);
        SysFreeString(name);
        return true;
    }

    IUIAutomationValuePattern* valuePattern = NULL;
    hr = element->lpVtbl->GetCurrentPattern(element, UIA_ValuePatternId, (IUnknown**)&valuePattern);

    if (SUCCEEDED(hr) && valuePattern) {
        BSTR value = NULL;
        hr = valuePattern->lpVtbl->get_CurrentValue(valuePattern, &value);
        if (SUCCEEDED(hr) && value) {
            WideCharToMultiByte(CP_UTF8, 0, value, -1, text, (int)text_size, NULL, NULL);
            SysFreeString(value);
            valuePattern->lpVtbl->Release(valuePattern);
            return true;
        }
        valuePattern->lpVtbl->Release(valuePattern);
    }

    return false;
}

static bool get_element_rect(IUIAutomationElement* element, RECT* rect) {
    if (!element || !rect) return false;

    HRESULT hr = element->lpVtbl->get_CurrentBoundingRectangle(element, rect);
    return SUCCEEDED(hr);
}



HWND winctrl_get_main_window(WinControlContext* ctx) {
    printf("Searching for main window of process %lu\n", ctx->current_process_id);

    struct EnumData {
        DWORD process_id;
        HWND window;
    } data = { ctx->current_process_id, NULL };

    EnumWindows(enum_windows_callback, (LPARAM)&data);

    if (!data.window) {
        printf("No suitable window found for process %lu\n", ctx->current_process_id);
    } else {
        printf("Found main window for process %lu\n", ctx->current_process_id);
    }

    return data.window;
}
bool winctrl_attach_process(WinControlContext* ctx, const char* process_name) {
    printf("Trying to attach to process: %s\n", process_name);

    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snapshot == INVALID_HANDLE_VALUE) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Failed to create process snapshot");
        return false;
    }

    PROCESSENTRY32W pe32;
    pe32.dwSize = sizeof(pe32);
    bool found = false;

    if (Process32FirstW(snapshot, &pe32)) {
        do {
            char curr_name[MAX_PATH];
            wcstombs_s(NULL, curr_name, sizeof(curr_name), pe32.szExeFile, _TRUNCATE);

            if (_stricmp(curr_name, process_name) == 0) {
                ctx->current_process_id = pe32.th32ProcessID;
                printf("Found process '%s' with PID: %lu\n", process_name, ctx->current_process_id);
                found = true;
                break;
            }
        } while (Process32NextW(snapshot, &pe32));
    }

    CloseHandle(snapshot);

    if (!found) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Process '%s' not found", process_name);
        return false;
    }

    ctx->current_window = winctrl_get_main_window(ctx);
    if (!ctx->current_window) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Could not find main window for process '%s'", process_name);
        return false;
    }

    return true;
}

bool winctrl_click_element(IUIAutomationElement* element) {
    if (!element) {
        return false;
    }

    RECT rect;
    if (!get_element_rect(element, &rect)) {
        return false;
    }

    int centerX = (rect.left + rect.right) / 2;
    int centerY = (rect.top + rect.bottom) / 2;

    printf("Clicking element at coordinates: %d, %d\n", centerX, centerY);

    BOOL isOffscreen = FALSE;
    element->lpVtbl->get_CurrentIsOffscreen(element, &isOffscreen);
    if (isOffscreen) {
        printf("Warning: Element appears to be offscreen\n");
        return false;
    }

    winctrl_click(centerX, centerY);
    Sleep(100);

    return true;
}

bool winctrl_is_element_enabled(IUIAutomationElement* element, bool* enabled) {
    if (!element || !enabled) return false;

    BOOL is_enabled = FALSE;
    HRESULT hr = element->lpVtbl->get_CurrentIsEnabled(element, &is_enabled);
    if (SUCCEEDED(hr)) {
        *enabled = is_enabled ? true : false;
        return true;
    }
    return false;
}

void winctrl_click(int x, int y) {
    SetCursorPos(x, y);
    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
    Sleep(10);
    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
}

void winctrl_right_click_coordinates(int x, int y) {
    SetCursorPos(x, y);
    mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
    Sleep(10);
    mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
}

void winctrl_double_click_coordinates(int x, int y) {
    SetCursorPos(x, y);
    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    Sleep(10);
    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
}

void winctrl_right_click(int x, int y) {
    SetCursorPos(x, y);
    mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
    Sleep(10);
    mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
}

void winctrl_double_click(int x, int y) {
    winctrl_click(x, y);
    Sleep(10);
    winctrl_click(x, y);
}

void winctrl_send_keys(WinControlContext* ctx, const char* text) {
    HKL layout = GetKeyboardLayout(0);

    while (*text) {
        SHORT vkey = VkKeyScanEx(*text, layout);
        BYTE scanCode = MapVirtualKeyEx(LOBYTE(vkey), 0, layout);

        keybd_event(LOBYTE(vkey), scanCode, 0, 0);
        Sleep(10);
        keybd_event(LOBYTE(vkey), scanCode, KEYEVENTF_KEYUP, 0);

        if (ctx->typing_delay_ms > 0) {
            Sleep(ctx->typing_delay_ms);
        }
        text++;
    }
}

void winctrl_sleep(int milliseconds) {
    Sleep(milliseconds);
}

bool winctrl_click_menu_item(WinControlContext* ctx, const char* menu, const char* item) {
    printf("Attempting to click menu '%s' and item '%s'\n", menu, item);

    IUIAutomationElement* menuElement = NULL;
    if (!winctrl_find_element_by_name(ctx, menu, &menuElement)) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Could not find menu: %s", menu);
        return false;
    }

    winctrl_click_element(menuElement);
    menuElement->lpVtbl->Release(menuElement);

    Sleep(500);

    IUIAutomationElement* itemElement = NULL;
    if (!winctrl_find_element_by_name(ctx, item, &itemElement)) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Could not find menu item: %s", item);
        return false;
    }

    bool result = winctrl_click_element(itemElement);
    itemElement->lpVtbl->Release(itemElement);

    return result;
}

bool winctrl_start_logging(WinControlContext* ctx, const char* base_filename) {
    if (!ctx || !base_filename) {
        if (ctx) sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Invalid parameters");
        return false;
    }

    time_t now;
    struct tm timeinfo;
    char timestamp[32];

    time(&now);
    if (localtime_s(&timeinfo, &now) != 0) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to get local time");
        return false;
    }

    if (strftime(timestamp, sizeof(timestamp), "%Y%m%d_%H%M%S", &timeinfo) == 0) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to format timestamp");
        return false;
    }

    if (snprintf(ctx->log_filename, sizeof(ctx->log_filename),
                "%s_%s.html", base_filename, timestamp) < 0) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to create filename");
        return false;
    }

    FILE* test;
    errno_t err = fopen_s(&test, ctx->log_filename, "r");
    if (err == 0) {
        fclose(test);
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Log file already exists: %s", ctx->log_filename);
        return false;
    }

    err = fopen_s(&ctx->log_file, ctx->log_filename, "w");
    if (err != 0) {
        char error_msg[256];
        strerror_s(error_msg, sizeof(error_msg), err);
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Failed to create log file: %s (%s)", ctx->log_filename, error_msg);
        return false;
    }

    if (fprintf(ctx->log_file,
        "<!DOCTYPE html>\n"
        "<html>\n"
        "<head>\n"
        "<title>Automation Log - %s</title>\n"
        "<style>\n"
        ".normal { color: black; }\n"
        ".warning { color: orange; }\n"
        ".error { color: red; }\n"
        ".header { font-size: 1.5em; font-weight: bold; }\n"
        "</style>\n"
        "</head>\n"
        "<body>\n"
        "<h1>Automation Log - Started at %s</h1>\n",
        timestamp, timestamp) < 0) {

        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to write HTML header");
        fclose(ctx->log_file);
        ctx->log_file = NULL;
        return false;
    }

    if (fflush(ctx->log_file) != 0) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to flush log file");
        fclose(ctx->log_file);
        ctx->log_file = NULL;
        return false;
    }

    return true;
}

void winctrl_log(WinControlContext* ctx, LogLevel level, const char* message) {
    if (!ctx || !ctx->log_file || !message) {
        return;
    }

    time_t now;
    struct tm timeinfo;
    char timestamp[32];

    time(&now);
    if (localtime_s(&timeinfo, &now) != 0) {
        return;
    }

    if (strftime(timestamp, sizeof(timestamp), "%H:%M:%S", &timeinfo) == 0) {
        return;
    }

    const char* css_class;
    switch (level) {
        case LOG_WARNING: css_class = "warning"; break;
        case LOG_ERROR: css_class = "error"; break;
        case LOG_HEADER: css_class = "header"; break;
        default: css_class = "normal";
    }

    if (fprintf(ctx->log_file,
        "<p class=\"%s\">[%s] %s</p>\n",
        css_class, timestamp, message) < 0) {
        return;
    }

    fflush(ctx->log_file);
}

void winctrl_end_logging(WinControlContext* ctx) {
    if (!ctx || !ctx->log_file) {
        return;
    }

    fprintf(ctx->log_file,
        "<hr>\n"
        "<p>Log ended at: %s</p>\n"
        "</body>\n"
        "</html>\n",
        ctx->log_filename);

    fclose(ctx->log_file);
    ctx->log_file = NULL;
    ctx->log_filename[0] = '\0';
}

bool winctrl_attach_pid(WinControlContext* ctx, DWORD process_id) {
    if (!winctrl_is_process_running(process_id)) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Process ID %lu not found", process_id);
        return false;
    }

    ctx->current_process_id = process_id;
    ctx->current_window = winctrl_get_main_window(ctx);

    if (!ctx->current_window) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "Could not find main window for process ID %lu", process_id);
        return false;
    }

    return true;
}

void winctrl_send_keys_with_modifier(WinModifierKeys modifiers, WORD key) {
    if (modifiers & WMOD_CTRL) {
        keybd_event(VK_CONTROL, 0, 0, 0);
    }
    if (modifiers & WMOD_ALT) {
        keybd_event(VK_MENU, 0, 0, 0);
    }
    if (modifiers & WMOD_SHIFT) {
        keybd_event(VK_SHIFT, 0, 0, 0);
    }
    if (modifiers & WMOD_WIN) {
        keybd_event(VK_LWIN, 0, 0, 0);
    }

    keybd_event(key, 0, 0, 0);
    Sleep(10);
    keybd_event(key, 0, KEYEVENTF_KEYUP, 0);

    if (modifiers & WMOD_WIN) {
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0);
    }
    if (modifiers & WMOD_SHIFT) {
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
    }
    if (modifiers & WMOD_ALT) {
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
    }
    if (modifiers & WMOD_CTRL) {
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
    }
}

bool winctrl_get_element_text_by_properties(WinControlContext* ctx,
    const ElementProperties* props,
    char* text_out,
    size_t text_out_size) {

    IUIAutomationElement* element = NULL;
    if (!winctrl_find_element_by_properties(ctx, props, &element)) {
        return false;
    }

    BSTR bstr_value = NULL;
    HRESULT hr = element->lpVtbl->get_CurrentName(element, &bstr_value);

    bool success = false;
    if (SUCCEEDED(hr) && bstr_value) {
        WideCharToMultiByte(CP_UTF8, 0, bstr_value, -1, text_out, (int)text_out_size, NULL, NULL);
        success = true;
    }

    if (bstr_value) {
        SysFreeString(bstr_value);
    }
    element->lpVtbl->Release(element);

    return success;
}

bool winctrl_is_process_running(DWORD process_id) {
    HANDLE process = OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, process_id);
    if (process == NULL) {
        return false;
    }

    DWORD exit_code;
    bool running = GetExitCodeProcess(process, &exit_code) && exit_code == STILL_ACTIVE;
    CloseHandle(process);
    return running;
}
bool winctrl_bring_to_front(WinControlContext* ctx) {
    if (!ctx->current_window) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error),
            "No window attached");
        return false;
    }

    if (IsIconic(ctx->current_window)) {
        ShowWindow(ctx->current_window, SW_RESTORE);
    }
    return SetForegroundWindow(ctx->current_window) != 0;
}

const char* winctrl_get_last_error(WinControlContext* ctx) {
    return ctx->last_error;
}

const char* winctrl_get_variable(WinControlContext* ctx, const char* name) {
    for (int i = 0; i < ctx->vars.variable_count; i++) {
        if (strcmp(ctx->vars.variables[i].name, name) == 0) {
            return ctx->vars.variables[i].value;
        }
    }
    return NULL;
}

static bool handle_click(WinControlContext* ctx, const Command* cmd) {
    winctrl_click(atoi(cmd->params[0]), atoi(cmd->params[1]));
    return true;
}

static bool handle_send_keystroke(WinControlContext* ctx, const Command* cmd) {
    const char* text_to_send;
    if (cmd->params[0][0] == '$') {
        text_to_send = winctrl_get_variable(ctx, cmd->params[0] + 1);
        if (!text_to_send) {
            sprintf_s(ctx->last_error, sizeof(ctx->last_error),
                "Variable not found: %s", cmd->params[0] + 1);
            return false;
        }
    } else {
        text_to_send = cmd->params[0];
    }
    printf("Sending keystroke: %s\n", text_to_send);
    winctrl_send_keys(ctx, text_to_send);
    return true;
}

static bool handle_start_log(WinControlContext* ctx, const Command* cmd) {
    return winctrl_start_logging(ctx, cmd->params[0]);
}

static bool handle_log(WinControlContext* ctx, const Command* cmd) {
    winctrl_log(ctx, LOG_NORMAL, cmd->params[0]);
    return true;
}

static bool handle_log_warning(WinControlContext* ctx, const Command* cmd) {
    winctrl_log(ctx, LOG_WARNING, cmd->params[0]);
    return true;
}

static bool handle_log_error(WinControlContext* ctx, const Command* cmd) {
    winctrl_log(ctx, LOG_ERROR, cmd->params[0]);
    return true;
}

static bool handle_log_header(WinControlContext* ctx, const Command* cmd) {
    winctrl_log(ctx, LOG_HEADER, cmd->params[0]);
    return true;
}

static bool handle_end_log(WinControlContext* ctx, const Command* cmd) {
    winctrl_end_logging(ctx);
    return true;
}

static bool handle_sleep(WinControlContext* ctx, const Command* cmd) {
    int ms = atoi(cmd->params[0]);
    printf("Sleeping for %d ms\n", ms);
    winctrl_sleep(ms);
    return true;
}

static bool handle_attach_process(WinControlContext* ctx, const Command* cmd) {
    printf("Attaching to process: %s\n", cmd->params[0]);
    return winctrl_attach_process(ctx, cmd->params[0]);
}

static bool handle_bring_to_front(WinControlContext* ctx, const Command* cmd) {
    printf("Bringing window to front\n");
    return winctrl_bring_to_front(ctx);
}

static bool handle_right_click(WinControlContext* ctx, const Command* cmd) {
    winctrl_right_click_coordinates(atoi(cmd->params[0]), atoi(cmd->params[1]));
    return true;
}

static bool handle_double_click(WinControlContext* ctx, const Command* cmd) {
    winctrl_double_click_coordinates(atoi(cmd->params[0]), atoi(cmd->params[1]));
    return true;
}

static bool handle_contains_element_text(WinControlContext* ctx, const Command* cmd) {
    ElementProperties props;
    props.automation_id = cmd->params[0];
    props.class_name = cmd->params[1];
    props.control_type = atoi(cmd->params[2]);

    char element_text[256] = {0};
    if (!winctrl_get_element_text_by_properties(ctx, &props, element_text, sizeof(element_text))) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Could not get element text");
        return false;
    }

    const char* search_text;
    if (cmd->params[3][0] == '$') {
        search_text = winctrl_get_variable(ctx, cmd->params[3] + 1);
        if (!search_text) {
            sprintf_s(ctx->last_error, sizeof(ctx->last_error),
                "Variable not found: %s", cmd->params[3] + 1);
            return false;
        }
    } else {
        search_text = cmd->params[3];
    }

    bool contains = (strstr(element_text, search_text) != NULL);
    winctrl_set_variable(ctx, "_CONTAINS_RESULT", contains ? "true" : "false");

    printf("Checking if element text '%s' contains '%s': %s\n",
        element_text, search_text, contains ? "yes" : "no");

    return true;
}

static bool handle_right_click_element(WinControlContext* ctx, const Command* cmd) {
    ElementProperties props;
    props.automation_id = strcmp(cmd->params[0], "null") == 0 ? NULL : cmd->params[0];
    props.class_name = strcmp(cmd->params[1], "null") == 0 ? NULL : cmd->params[1];
    props.control_type = atoi(cmd->params[2]);

    IUIAutomationElement* element = NULL;
    if (winctrl_find_element_by_properties(ctx, &props, &element)) {
        printf("Found element, right-clicking...\n");
        bool clicked = winctrl_right_click_element(element);
        element->lpVtbl->Release(element);
        return clicked;
    }
    return false;
}

static bool handle_double_click_element(WinControlContext* ctx, const Command* cmd) {
    ElementProperties props;
    props.automation_id = strcmp(cmd->params[0], "null") == 0 ? NULL : cmd->params[0];
    props.class_name = strcmp(cmd->params[1], "null") == 0 ? NULL : cmd->params[1];
    props.control_type = atoi(cmd->params[2]);

    IUIAutomationElement* element = NULL;
    if (winctrl_find_element_by_properties(ctx, &props, &element)) {
        printf("Found element, double-clicking...\n");
        bool clicked = winctrl_double_click_element(element);
        element->lpVtbl->Release(element);
        return clicked;
    }
    return false;
}

static bool handle_send_mod_key(WinControlContext* ctx, const Command* cmd) {
    WinModifierKeys mod = WMOD_NONE;
    WORD key = 0;

    if (strcmp(cmd->params[0], "CTRL") == 0) mod = WMOD_CTRL;
    else if (strcmp(cmd->params[0], "ALT") == 0) mod = WMOD_ALT;
    else if (strcmp(cmd->params[0], "SHIFT") == 0) mod = WMOD_SHIFT;
    else if (strcmp(cmd->params[0], "WIN") == 0) mod = WMOD_WIN;

    if (strlen(cmd->params[1]) == 1) {
        key = VkKeyScanEx(cmd->params[1][0], GetKeyboardLayout(0)) & 0xFF;
    } else {
        if (strcmp(cmd->params[1], "TAB") == 0) key = VK_TAB;
        else if (strcmp(cmd->params[1], "ENTER") == 0) key = VK_RETURN;
        else if (strcmp(cmd->params[1], "ESC") == 0) key = VK_ESCAPE;
    }

    if (key != 0) {
        printf("Sending modified key: %s + %s\n", cmd->params[0], cmd->params[1]);
        winctrl_send_keys_with_modifier(mod, key);
        return true;
    }

    sprintf_s(ctx->last_error, sizeof(ctx->last_error),
        "Invalid key combination: %s + %s", cmd->params[0], cmd->params[1]);
    return false;
}

static bool handle_set(WinControlContext* ctx, const Command* cmd) {
    return winctrl_set_variable(ctx, cmd->params[0], cmd->params[1]);
}

static bool handle_if(WinControlContext* ctx, const Command* cmd) {
    char condition[512] = {0};
    strncpy_s(condition, sizeof(condition), cmd->params[0], _TRUNCATE);
    for (int i = 1; i < cmd->param_count; i++) {
        strncat_s(condition, sizeof(condition), " ", _TRUNCATE);
        strncat_s(condition, sizeof(condition), cmd->params[i], _TRUNCATE);
    }
    printf("Evaluating condition: %s\n", condition);
    bool condition_met = evaluate_condition(ctx, condition);
    winctrl_set_variable(ctx, "_IF_CONDITION", condition_met ? "true" : "false");
    return true;
}

static bool handle_endif(WinControlContext* ctx, const Command* cmd) {
    winctrl_set_variable(ctx, "_IF_CONDITION", "true");
    return true;
}

static bool handle_set_delay(WinControlContext* ctx, const Command* cmd) {
    ctx->typing_delay_ms = atoi(cmd->params[0]);
    return true;
}

static bool handle_send_multi_mod_key(WinControlContext* ctx, const Command* cmd) {
    WinModifierKeys mods = WMOD_NONE;
    for (int i = 0; i < cmd->param_count - 1; i++) {
        if (strcmp(cmd->params[i], "CTRL") == 0) mods |= WMOD_CTRL;
        else if (strcmp(cmd->params[i], "ALT") == 0) mods |= WMOD_ALT;
        else if (strcmp(cmd->params[i], "SHIFT") == 0) mods |= WMOD_SHIFT;
        else if (strcmp(cmd->params[i], "WIN") == 0) mods |= WMOD_WIN;
    }

    WORD key = 0;
    const char* keyStr = cmd->params[cmd->param_count - 1];

    if (strlen(keyStr) == 1) {
        key = VkKeyScanEx(keyStr[0], GetKeyboardLayout(0)) & 0xFF;
    } else {
        if (strcmp(keyStr, "TAB") == 0) key = VK_TAB;
        else if (strcmp(keyStr, "ENTER") == 0) key = VK_RETURN;
        else if (strcmp(keyStr, "ESC") == 0) key = VK_ESCAPE;
        else if (strcmp(keyStr, "DELETE") == 0) key = VK_DELETE;
    }

    if (key != 0) {
        winctrl_send_keys_with_modifier(mods, key);
        return true;
    }

    sprintf_s(ctx->last_error, sizeof(ctx->last_error),
        "Invalid key combination. Modifiers: %d, Key: %s", mods, keyStr);
    return false;
}

static bool handle_click_element(WinControlContext* ctx, const Command* cmd) {
    ElementProperties props;
    props.automation_id = strcmp(cmd->params[0], "null") == 0 ? NULL : cmd->params[0];
    props.class_name = strcmp(cmd->params[1], "null") == 0 ? NULL : cmd->params[1];
    props.control_type = atoi(cmd->params[2]);

    IUIAutomationElement* element = NULL;
    if (winctrl_find_element_by_properties(ctx, &props, &element)) {
        printf("Found element, clicking...\n");
        bool clicked = winctrl_click_element(element);
        element->lpVtbl->Release(element);
        return clicked;
    }

    sprintf_s(ctx->last_error, sizeof(ctx->last_error),
        "Could not find element with specified properties");
    return false;
}

typedef struct {
    const char* name;
    int param_count;
    CommandHandler handler;
} CommandDefinition;

static const CommandDefinition COMMAND_TABLE[] = {
    {"Click", 2, handle_click},
    {"SendKeystroke", 1, handle_send_keystroke},
    {"StartLog", 1, handle_start_log},
    {"Log", 1, handle_log},
    {"LogWarning", 1, handle_log_warning},
    {"LogError", 1, handle_log_error},
    {"LogHeader", 1, handle_log_header},
    {"EndLog", 0, handle_end_log},
    {"Sleep", 1, handle_sleep},
    {"AttachProcess", 1, handle_attach_process},
    {"BringToFront", 0, handle_bring_to_front},
    {"RightClick", 2, handle_right_click},
    {"DoubleClick", 2, handle_double_click},
    {"ContainsElementText", 4, handle_contains_element_text},
    {"RightClickElementByProperties", 3, handle_right_click_element},
    {"DoubleClickElementByProperties", 3, handle_double_click_element},
    {"SendModKey", 2, handle_send_mod_key},
    {"SET", 2, handle_set},
    {"IF", -1, handle_if},
    {"ENDIF", 0, handle_endif},
    {"SetDelay", 1, handle_set_delay},
    {"SendMultiModKey", -1, handle_send_multi_mod_key},
    {"ClickElementByProperties", 3, handle_click_element},
    {NULL, 0, NULL}
};

bool winctrl_execute_command(WinControlContext* ctx, const Command* cmd) {
    printf("Executing command: %s with %d parameters\n", cmd->name, cmd->param_count);
    for(int i = 0; i < cmd->param_count; i++) {
        printf("Parameter %d: '%s'\n", i, cmd->params[i]);
    }

    for (const CommandDefinition* def = COMMAND_TABLE; def->name != NULL; def++) {
        if (strcmp(cmd->name, def->name) == 0) {
            if (def->param_count != -1 && cmd->param_count != def->param_count) {
                sprintf_s(ctx->last_error, sizeof(ctx->last_error),
                    "Invalid parameter count for %s: expected %d, got %d",
                    cmd->name, def->param_count, cmd->param_count);
                return false;
            }
            return def->handler(ctx, cmd);
        }
    }

    sprintf_s(ctx->last_error, sizeof(ctx->last_error),
        "Unknown command: %s", cmd->name);
    return false;
}

static bool create_property_condition(WinControlContext* ctx,
                                    int property_id,
                                    VARIANT value,
                                    IUIAutomationCondition** condition) {
    HRESULT hr = ctx->automation->lpVtbl->CreatePropertyCondition(
        ctx->automation,
        property_id,
        value,
        condition
    );
    return SUCCEEDED(hr);
}


bool winctrl_find_element_by_properties(WinControlContext* ctx,
    const ElementProperties* props,
    IUIAutomationElement** element) {

    if (!ctx->current_window) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "No window attached");
        return false;
    }

    printf("Looking for element with properties...\n");

    IUIAutomationCondition* conditions[3] = {NULL};
    int condition_count = 0;

    if (props->automation_id) {
        WCHAR wide_id[256];
        MultiByteToWideChar(CP_UTF8, 0, props->automation_id, -1, wide_id, 256);
        VARIANT var;
        var.vt = VT_BSTR;
        var.bstrVal = SysAllocString(wide_id);
        ctx->automation->lpVtbl->CreatePropertyCondition(
            ctx->automation,
            UIA_AutomationIdPropertyId,
            var,
            &conditions[condition_count++]
        );
        SysFreeString(var.bstrVal);
    }

    if (props->class_name) {
        WCHAR wide_class[256];
        MultiByteToWideChar(CP_UTF8, 0, props->class_name, -1, wide_class, 256);
        VARIANT var;
        var.vt = VT_BSTR;
        var.bstrVal = SysAllocString(wide_class);
        ctx->automation->lpVtbl->CreatePropertyCondition(
            ctx->automation,
            UIA_ClassNamePropertyId,
            var,
            &conditions[condition_count++]
        );
        SysFreeString(var.bstrVal);
    }

    if (props->control_type != -1) {
        VARIANT var;
        var.vt = VT_I4;
        var.lVal = props->control_type;
        ctx->automation->lpVtbl->CreatePropertyCondition(
            ctx->automation,
            UIA_ControlTypePropertyId,
            var,
            &conditions[condition_count++]
        );
    }

    IUIAutomationCondition* final_condition = NULL;
    if (condition_count > 0) {
        if (condition_count == 1) {
            final_condition = conditions[0];
        } else {
            final_condition = conditions[0];
            for (int i = 1; i < condition_count; i++) {
                IUIAutomationCondition* temp = final_condition;
                ctx->automation->lpVtbl->CreateAndCondition(
                    ctx->automation,
                    temp,
                    conditions[i],
                    &final_condition
                );
                if (i > 1) {
                    temp->lpVtbl->Release(temp);
                }
                conditions[i]->lpVtbl->Release(conditions[i]);
            }
            conditions[0]->lpVtbl->Release(conditions[0]);
        }
    }

    IUIAutomationElement* root = NULL;
    HRESULT hr = ctx->automation->lpVtbl->ElementFromHandle(
        ctx->automation,
        ctx->current_window,
        &root
    );

    if (SUCCEEDED(hr) && root) {
        hr = root->lpVtbl->FindFirst(
            root,
            TreeScope_Descendants,
            final_condition,
            element
        );
        root->lpVtbl->Release(root);
    }

    if (final_condition) {
        final_condition->lpVtbl->Release(final_condition);
    }

    if (SUCCEEDED(hr) && *element) {
        printf("Element found!\n");
        return true;
    }

    printf("Element not found (hr = 0x%lx)\n", hr);
    return false;
}


bool winctrl_find_element_by_name(WinControlContext* ctx, const char* name, IUIAutomationElement** element) {
    if (!ctx->current_window) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "No window attached");
        return false;
    }

    printf("Looking for element: %s\n", name);

    WCHAR wide_name[256];
    MultiByteToWideChar(CP_UTF8, 0, name, -1, wide_name, 256);

    VARIANT var;
    var.vt = VT_BSTR;
    var.bstrVal = SysAllocString(wide_name);

    IUIAutomationCondition* condition = NULL;
    HRESULT hr = ctx->automation->lpVtbl->CreatePropertyCondition(
        ctx->automation,
        UIA_NamePropertyId,
        var,
        &condition
    );

    if (FAILED(hr)) {
        SysFreeString(var.bstrVal);
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to create search condition");
        return false;
    }

    IUIAutomationElement* root = NULL;
    hr = ctx->automation->lpVtbl->ElementFromHandle(
        ctx->automation,
        ctx->current_window,
        &root
    );

    if (FAILED(hr)) {
        condition->lpVtbl->Release(condition);
        SysFreeString(var.bstrVal);
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Failed to get root element");
        return false;
    }

    hr = root->lpVtbl->FindFirst(
        root,
        TreeScope_Descendants,
        condition,
        element
    );

    root->lpVtbl->Release(root);
    condition->lpVtbl->Release(condition);
    SysFreeString(var.bstrVal);

    if (FAILED(hr) || !*element) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "Element not found: %s", name);
        return false;
    }

    printf("Found element: %s\n", name);
    return true;
}

bool winctrl_find_element_by_id(WinControlContext* ctx, const char* automation_id, IUIAutomationElement** element) {
    if (!ctx->current_window) {
        sprintf_s(ctx->last_error, sizeof(ctx->last_error), "No window attached");
        return false;
    }

    WCHAR wide_id[256];
    MultiByteToWideChar(CP_UTF8, 0, automation_id, -1, wide_id, 256);

    VARIANT var;
    var.vt = VT_BSTR;
    var.bstrVal = SysAllocString(wide_id);

    IUIAutomationCondition* condition = NULL;
    if (!create_property_condition(ctx, UIA_AutomationIdPropertyId, var, &condition)) {
        SysFreeString(var.bstrVal);
        return false;
    }

    IUIAutomationElement* root = NULL;
    ctx->automation->lpVtbl->ElementFromHandle(ctx->automation, ctx->current_window, &root);

    HRESULT hr = root->lpVtbl->FindFirst(root, TreeScope_Descendants, condition, element);

    root->lpVtbl->Release(root);
    condition->lpVtbl->Release(condition);
    SysFreeString(var.bstrVal);

    return SUCCEEDED(hr) && *element != NULL;
}

bool winctrl_wait_for_element(WinControlContext* ctx, const char* name, int timeout_ms, IUIAutomationElement** element) {
    int elapsed = 0;
    const int sleep_interval = 100;

    while (elapsed < timeout_ms) {
        if (winctrl_find_element_by_name(ctx, name, element)) {
            return true;
        }
        Sleep(sleep_interval);
        elapsed += sleep_interval;
    }

    sprintf_s(ctx->last_error, sizeof(ctx->last_error),
        "Timeout waiting for element: %s", name);
    return false;
}

bool winctrl_right_click_element(IUIAutomationElement* element) {
    RECT rect;
    if (!get_element_rect(element, &rect)) {
        return false;
    }

    int centerX = (rect.left + rect.right) / 2;
    int centerY = (rect.top + rect.bottom) / 2;

    printf("Right-clicking element at coordinates: %d, %d\n", centerX, centerY);

    SetCursorPos(centerX, centerY);
    mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
    Sleep(10);
    mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
    Sleep(100);

    return true;
}



bool winctrl_double_click_element(IUIAutomationElement* element) {
    RECT rect;
    if (!get_element_rect(element, &rect)) {
        return false;
    }

    int centerX = (rect.left + rect.right) / 2;
    int centerY = (rect.top + rect.bottom) / 2;

    printf("Double-clicking element at coordinates: %d, %d\n", centerX, centerY);

    SetCursorPos(centerX, centerY);
    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    Sleep(10);
    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    Sleep(100);

    return true;
}

bool evaluate_condition(WinControlContext* ctx, const char* condition) {
    if (strncmp(condition, "ElementExists", 12) == 0) {
        const char* element_name = condition + 13;
        IUIAutomationElement* element = NULL;
        ElementProperties props = {0};
        props.automation_id = element_name;
        return winctrl_find_element_by_properties(ctx, &props, &element);
    }

    if (strncmp(condition, "ElementNotExists", 15) == 0) {
        const char* element_name = condition + 16;
        IUIAutomationElement* element = NULL;
        ElementProperties props = {0};
        props.automation_id = element_name;
        bool exists = winctrl_find_element_by_properties(ctx, &props, &element);
        if (element) element->lpVtbl->Release(element);
        return !exists;
    }

    if (strncmp(condition, "ContainsElementText", 18) == 0) {
        char params[4][256] = {0};
        int param_count = 0;

        const char* start = condition + 19;
        char* token;
        char* next_token = NULL;
        char condition_copy[512];
        strncpy_s(condition_copy, sizeof(condition_copy), start, _TRUNCATE);

        token = strtok_s(condition_copy, " ", &next_token);
        while (token && param_count < 4) {
            if (token[0] == '"') {
                token++;
                char* end_quote = strrchr(token, '"');
                if (end_quote) *end_quote = '\0';
            }
            strncpy_s(params[param_count], sizeof(params[param_count]), token, _TRUNCATE);
            param_count++;
            token = strtok_s(NULL, " ", &next_token);
        }

        if (param_count == 4) {
            ElementProperties props = {0};
            props.automation_id = params[0];
            props.class_name = params[1];
            props.control_type = atoi(params[2]);

            char element_text[256] = {0};
            IUIAutomationElement* element = NULL;

            const char* compare_text;
            if (params[3][0] == '$') {
                compare_text = winctrl_get_variable(ctx, params[3] + 1);
                if (!compare_text) {
                    return false;
                }
            } else {
                compare_text = params[3];
            }

            if (winctrl_find_element_by_properties(ctx, &props, &element)) {
                BSTR bstr_value = NULL;
                if (SUCCEEDED(element->lpVtbl->get_CurrentName(element, &bstr_value)) && bstr_value) {
                    WideCharToMultiByte(CP_UTF8, 0, bstr_value, -1, element_text, sizeof(element_text), NULL, NULL);
                    SysFreeString(bstr_value);
                    element->lpVtbl->Release(element);

                    return (strstr(element_text, compare_text) != NULL);
                }
                element->lpVtbl->Release(element);
            }
        }
        return false;
    }

    return false;
}



int winctrl_parse_script(const char* filename, Command* commands, int max_commands) {
    FILE* file;
    if (fopen_s(&file, filename, "r") != 0) {
        return -1;
    }

    char line[512];
    int count = 0;
    bool skip_commands = false;

    while (fgets(line, sizeof(line), file) && count < max_commands) {
        char* comment = strchr(line, '#');
        if (comment) {
            *comment = '\0';
        }

        bool isEmpty = true;
        for (char* p = line; *p; p++) {
            if (!isspace(*p)) {
                isEmpty = false;
                break;
            }
        }
        if (isEmpty) continue;

        Command* current_cmd = &commands[count];
        current_cmd->param_count = 0;

        char* next_token = NULL;
        char* token = strtok_s(line, " \t\n\r", &next_token);
        if (!token) continue;

        strncpy_s(current_cmd->name, sizeof(current_cmd->name), token, _TRUNCATE);

        while ((token = strtok_s(NULL, " \t\n\r", &next_token)) != NULL &&
               current_cmd->param_count < 4) {

            if (token[0] == '"') {
                token++;

                char* end_quote = strrchr(token, '"');
                if (end_quote) {
                    *end_quote = '\0';
                    strncpy_s(current_cmd->params[current_cmd->param_count],
                        sizeof(current_cmd->params[current_cmd->param_count]),
                        token, _TRUNCATE);
                } else {
                    char full_param[256] = {0};
                    strncpy_s(full_param, sizeof(full_param), token, _TRUNCATE);

                    while ((token = strtok_s(NULL, "\n\r", &next_token)) != NULL) {
                        strncat_s(full_param, sizeof(full_param), " ", _TRUNCATE);
                        strncat_s(full_param, sizeof(full_param), token, _TRUNCATE);

                        if (strchr(token, '"')) {
                            char* end = strrchr(full_param, '"');
                            if (end) *end = '\0';
                            break;
                        }
                    }

                    strncpy_s(current_cmd->params[current_cmd->param_count],
                        sizeof(current_cmd->params[current_cmd->param_count]),
                        full_param, _TRUNCATE);
                }
            } else {
                strncpy_s(current_cmd->params[current_cmd->param_count],
                    sizeof(current_cmd->params[current_cmd->param_count]),
                    token, _TRUNCATE);
            }

            current_cmd->param_count++;
        }

        count++;
    }

    fclose(file);
    return count;
}