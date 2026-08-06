#define _CRT_SECURE_NO_WARNINGS
#pragma warning(disable: 4819)

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

// 데이터를 파싱하고 CSV 파일에 이어쓰는 함수
void process_and_save_data(const char* raw_data, const char* output_filename) {
    int yy, mm, dd, hh, min, ss;
    float temp, viscosity, percent;
    int rotor;
    float speed;

    int match_count = sscanf(raw_data, "%d , %d , %d , %d , %d , %d , %f , %d , %f , %f , %f",
        &yy, &mm, &dd, &hh, &min, &ss,
        &temp, &rotor, &speed, &viscosity, &percent);

    if (match_count == 11) {
        FILE* out_file = fopen(output_filename, "a");
        if (out_file != NULL) {
            fseek(out_file, 0, SEEK_END);
            if (ftell(out_file) == 0) {
                fprintf(out_file, "Date,Time,Temperature(C),Rotor,Speed(RPM),Viscosity(mPa.s),Percent(%%)\n");
            }

            fprintf(out_file, "20%02d-%02d-%02d,%02d:%02d:%02d,%.1f,%d,%.1f,%.2f,%.1f\n",
                yy, mm, dd, hh, min, ss, temp, rotor, speed, viscosity, percent);
            fclose(out_file);

            printf("[Saved] 20%02d-%02d-%02d %02d:%02d:%02d (Viscosity: %.2f mPa.s)\n",
                yy, mm, dd, hh, min, ss, viscosity);
        }
    }
    else {
        printf("[Format Error] Data: %s\n", raw_data);
    }
}

// 가장 높은 번호의 활성화된 COM 포트를 자동으로 찾아주는 함수
int find_auto_com_port() {
    int auto_port = -1;

    // COM1부터 COM256까지 스캔하면서 가장 높은 번호를 찾음
    for (int i = 1; i <= 256; i++) {
        char port_name[16];
        sprintf(port_name, "COM%d", i);
        char target_path[256];

        if (QueryDosDeviceA(port_name, target_path, 256) > 0) {
            auto_port = i; // 발견될 때마다 덮어씌워서 가장 높은 번호가 남게 함
        }
    }
    return auto_port;
}

int main() {
    const char* output_file = "viscometer_realtime_data.csv";

    printf("=== Real-time Viscometer Logger ===\n");
    printf("Scanning for available COM ports...\n");

    // 1. 포트 자동 검색
    int port_num = find_auto_com_port();

    if (port_num == -1) {
        printf("\nError: No COM ports found!\n");
        printf("-> Please check if the USB cable is properly connected.\n");
        system("pause");
        return 1;
    }

    printf("-> Auto-selected port: COM%d\n", port_num);

    char port_name[20];
    sprintf(port_name, "\\\\.\\COM%d", port_num);

    // 2. 시리얼 포트 열기
    HANDLE hSerial = CreateFileA(port_name, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
    if (hSerial == INVALID_HANDLE_VALUE) {
        printf("\nError: Cannot open COM%d\n", port_num);
        printf("-> Make sure no other program (like PuTTY) is using this port.\n");
        system("pause");
        return 1;
    }

    DCB dcbSerialParams = { 0 };
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
    if (!GetCommState(hSerial, &dcbSerialParams)) {
        printf("\nError: Failed to get COM state.\n");
        CloseHandle(hSerial);
        return 1;
    }

    dcbSerialParams.BaudRate = CBR_9600;
    dcbSerialParams.ByteSize = 8;
    dcbSerialParams.StopBits = ONESTOPBIT;
    dcbSerialParams.Parity = NOPARITY;

    if (!SetCommState(hSerial, &dcbSerialParams)) {
        printf("\nError: Failed to set COM state.\n");
        CloseHandle(hSerial);
        return 1;
    }

    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 50;
    timeouts.ReadTotalTimeoutConstant = 50;
    timeouts.ReadTotalTimeoutMultiplier = 10;
    SetCommTimeouts(hSerial, &timeouts);

    printf("Connected to COM%d successfully!\n", port_num);
    printf("Waiting for data... (Press Ctrl+C to stop)\n\n");

    char buffer[1024];
    int buf_idx = 0;
    int capturing = 0;

    // 3. 실시간 데이터 수신 및 처리 루프
    while (1) {
        char c;
        DWORD bytesRead;

        if (ReadFile(hSerial, &c, 1, &bytesRead, NULL) && bytesRead > 0) {
            if (c == '$') {
                capturing = 1;
                buf_idx = 0;
            }
            else if (c == '*' && capturing) {
                buffer[buf_idx] = '\0';
                process_and_save_data(buffer, output_file);
                capturing = 0;
            }
            else if (capturing) {
                if (buf_idx < sizeof(buffer) - 1) {
                    buffer[buf_idx++] = c;
                }
            }
        }
    }

    CloseHandle(hSerial);
    return 0;
}