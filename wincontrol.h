#ifndef WINCONTROL_H
#define WINCONTROL_H

#include <windows.h>
#include <stdbool.h>
#include <time.h>
#include <stdio.h>

typedef interface IUIAutomation IUIAutomation;
typedef interface IUIAutomationElement IUIAutomationElement;

#define MAX_VARIABLES 100
#define MAX_VAR_NAME 32
#define MAX_VAR_VALUE 256

typedef enum {
    WMOD_NONE = 0,
    WMOD_CTRL = 1,
    WMOD_ALT = 2,
    WMOD_SHIFT = 4,
    WMOD_WIN = 8
} WinModifierKeys;

typedef enum {
    LOG_NORMAL,
    LOG_WARNING,
    LOG_ERROR,
    LOG_HEADER
} LogLevel;

typedef struct {
    char name[MAX_VAR_NAME];
    char value[MAX_VAR_VALUE];
} Variable;

typedef struct {
    Variable variables[MAX_VARIABLES];
    int variable_count;
} VariableContext;

typedef struct {
    char name[32];
    char params[4][256];
    int param_count;
} Command;

typedef struct {
    const char* automation_id;
    const char* class_name;
    int control_type;
} ElementProperties;

typedef struct {
    IUIAutomation* automation;
    DWORD current_process_id;
    HWND current_window;
    char last_error[256];
    VariableContext vars;
    FILE* log_file;
    char log_filename[256];
    int typing_delay_ms;
} WinControlContext;

bool winctrl_initialize(WinControlContext* ctx);
void winctrl_cleanup(WinControlContext* ctx);
const char* winctrl_get_last_error(WinControlContext* ctx);

void winctrl_click(int x, int y);
void winctrl_right_click(int x, int y);
void winctrl_double_click(int x, int y);
void winctrl_right_click_coordinates(int x, int y);
void winctrl_double_click_coordinates(int x, int y);
void winctrl_send_keys(WinControlContext* ctx, const char* text);
void winctrl_send_keys_with_modifier(WinModifierKeys modifiers, WORD key);
void winctrl_sleep(int milliseconds);


bool winctrl_attach_process(WinControlContext* ctx, const char* process_name);
bool winctrl_attach_pid(WinControlContext* ctx, DWORD process_id);
bool winctrl_bring_to_front(WinControlContext* ctx);
bool winctrl_is_process_running(DWORD process_id);
bool winctrl_find_element_by_properties(WinControlContext* ctx, const ElementProperties* props, IUIAutomationElement** element);
bool winctrl_find_element_by_name(WinControlContext* ctx, const char* name, IUIAutomationElement** element);
bool winctrl_find_element_by_id(WinControlContext* ctx, const char* automation_id, IUIAutomationElement** element);
bool winctrl_find_element_by_class(WinControlContext* ctx, const char* class_name, IUIAutomationElement** element);
bool winctrl_find_element_by_type(WinControlContext* ctx, int control_type, IUIAutomationElement** element);
bool winctrl_wait_for_element(WinControlContext* ctx, const char* name, int timeout_ms, IUIAutomationElement** element);
bool winctrl_get_element_text(IUIAutomationElement* element, char* text, size_t text_size);
bool winctrl_get_element_text_by_properties(WinControlContext* ctx, const ElementProperties* props, char* text_out, size_t text_out_size);
bool winctrl_set_element_value(IUIAutomationElement* element, const char* value);
bool winctrl_is_element_enabled(IUIAutomationElement* element, bool* enabled);
bool winctrl_is_element_visible(IUIAutomationElement* element, bool* visible);
bool winctrl_click_element(IUIAutomationElement* element);
bool winctrl_right_click_element(IUIAutomationElement* element);
bool winctrl_double_click_element(IUIAutomationElement* element);
bool winctrl_select_combo_item(IUIAutomationElement* element, const char* item);
bool winctrl_check_checkbox(IUIAutomationElement* element, bool check);
bool winctrl_expand_collapse(IUIAutomationElement* element, bool expand);
bool winctrl_click_menu_item(WinControlContext* ctx, const char* menu, const char* item);
bool winctrl_compare_text(const char* text1, const char* text2);
bool winctrl_set_variable(WinControlContext* ctx, const char* name, const char* value);
const char* winctrl_get_variable(WinControlContext* ctx, const char* name);
bool evaluate_condition(WinControlContext* ctx, const char* condition);
bool winctrl_start_logging(WinControlContext* ctx, const char* base_filename);
void winctrl_log(WinControlContext* ctx, LogLevel level, const char* message);
void winctrl_end_logging(WinControlContext* ctx);
int winctrl_parse_script(const char* filename, Command* commands, int max_commands);
bool winctrl_execute_command(WinControlContext* ctx, const Command* cmd);

#endif