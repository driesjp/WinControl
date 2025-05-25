#include "wincontrol.h"
#include <stdio.h>

void print_usage(void) {
    printf("WinControl - Windows Automation Tool\n");
    printf("Copyright (C) 2024 Dries R. M. Swartele.\n\n");
    printf("This program is free software: you can redistribute it and/or modify it\n");
    printf("under the terms of the GNU General Public License as published by the Free\n");
    printf("Software Foundation, either version 3 of the License, or (at your option)\n");
    printf("any later version.\n\n");

    printf("This program is distributed in the hope that it will be useful,\n");
    printf("but WITHOUT ANY WARRANTY; without even the implied warranty of\n");
    printf("MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the\n");
    printf("GNU General Public License for more details.\n\n");

    printf("You should have received a copy of the GNU General Public License\n");
    printf("along with this program.  If not, see <https://www.gnu.org/licenses/>.\n\n\n");



    printf("More info: http://www.dries.jp\n\n");
    printf("Usage: WinControl.exe -s <script_file>\n\n");
    printf("Available script commands:\n");
    printf("  AttachProcess \"processname\" - Attach to a running process\n");
    printf("  BringToFront                  - Bring current window to front\n");
    printf("  Click x y                     - Click at coordinates\n");
    printf("  RightClick x y                - Right click at coordinates\n");
    printf("  DoubleClick x y               - Double click at coordinates\n");
    printf("  SendKeystroke \"text\"        - Send keystrokes\n");
    printf("  Sleep milliseconds            - Wait specified time\n\n");
    printf("  SET mytext \"Hello World\"    - Set variable\n  e.g.\n");
    printf("  SendKeystroke \"$mytext\"     - Use variable for SendKeyStroke\n\n");

    printf("  Conditional execution:\n");
    printf("  IF ElementExists \"id\" \"class\" \"type\"\n    # code\n  ENDIF\n\n");
    printf("  IF ElementNotExists \"id\" \"class\" \"type\"\n    # code\n  ENDIF\n\n");
    printf("  IF ContainsElementText \"textbox_id\" \"textbox_class\" \"50011\" \"$mytext\"\n    #blabla\n  ENDIF\n\n");


    printf("");

    printf("  Logging:\n");
    printf("  StartLog \"AutomationTest\"       - Create log file and open filestream\n");
    printf("  LogHeader \"Starting test\"       - Create big log entry\n");
    printf("  Log \"Normal message\"            - Create normal log entry\n");
    printf("  LogWarning \"Warning message\"    - Create warning log entry\n");
    printf("  LogError \"Error message\"        - Create error log entry\n");
    printf("  EndLog                            - Close filestream\n\n\n");




    printf("Example script:\n");
    printf("  AttachProcess \"notepad.exe\"\n");
    printf("  Sleep 1000\n");
    printf("  SendKeystroke \"Hello World\"\n");

}

int main(int argc, char* argv[]) {
    if (argc != 3 || strcmp(argv[1], "-s") != 0) {
        print_usage();
        return 1;
    }

    WinControlContext ctx = { 0 };
    if (!winctrl_initialize(&ctx)) {
        printf("Error: %s\n", winctrl_get_last_error(&ctx));
        return 1;
    }

    Command commands[100];
    int cmd_count = winctrl_parse_script(argv[2], commands, 100);

    if (cmd_count < 0) {
        printf("Error parsing script file: %s\n", winctrl_get_last_error(&ctx));
        winctrl_cleanup(&ctx);
        return 1;
    }

    printf("Executing script with %d commands...\n", cmd_count);

    for (int i = 0; i < cmd_count; i++) {
        if (!winctrl_execute_command(&ctx, &commands[i])) {
            printf("Error executing command '%s': %s\n",
                commands[i].name, winctrl_get_last_error(&ctx));
            break;
        }
    }

    char current_dir[MAX_PATH];
    GetCurrentDirectoryA(MAX_PATH, current_dir);
    printf("Current directory: %s\n", current_dir);
    printf("Script: %s\n", argv[2]);

    winctrl_cleanup(&ctx);
    return 0;
}