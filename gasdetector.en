
#include <windows.h>
#include <stdio.h>
#include <stdint.h>
#include <stdlib.h>

/* --- 화면 색상 설정을 위한 매크로 --- */
#define COLOR_DEFAULT 7
#define COLOR_GREEN   10
#define COLOR_CYAN    11
#define COLOR_RED     12
#define COLOR_YELLOW  14

/* 콘솔 글자 색상 변경 함수 */
void SetColor(int color)
{
    SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color);
}

/* 화면 깜빡임 방지를 위한 커서 이동 함수 (구형 C언어 표준으로 수정) */
void MoveCursor(int x, int y)
{
    COORD pos;
    pos.X = (SHORT)x;
    pos.Y = (SHORT)y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), pos);
}

/* ----------------------------
   CRC16 Modbus
---------------------------- */
uint16_t ModbusCRC16(uint8_t* data, uint16_t length)
{
    uint16_t crc = 0xFFFF;
    int pos, i;

    for (pos = 0; pos < length; pos++)
    {
        crc ^= data[pos];
        for (i = 0; i < 8; i++)
        {
            if (crc & 0x0001) crc = (crc >> 1) ^ 0xA001;
            else             crc >>= 1;
        }
    }
    return crc;
}

/* ----------------------------
   Serial Open
---------------------------- */
HANDLE OpenSerial(char* port)
{
    HANDLE hSerial;
    DCB dcb = { 0 };
    COMMTIMEOUTS timeout = { 0 };

    hSerial = CreateFileA(port, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (hSerial == INVALID_HANDLE_VALUE) return NULL;

    dcb.DCBlength = sizeof(DCB);
    if (!GetCommState(hSerial, &dcb)) { CloseHandle(hSerial); return NULL; }

    dcb.BaudRate = CBR_9600;
    dcb.ByteSize = 8;
    dcb.StopBits = ONESTOPBIT;
    dcb.Parity = NOPARITY;
    if (!SetCommState(hSerial, &dcb)) { CloseHandle(hSerial); return NULL; }

    timeout.ReadIntervalTimeout = 50;
    timeout.ReadTotalTimeoutConstant = 100;
    timeout.ReadTotalTimeoutMultiplier = 10;
    if (!SetCommTimeouts(hSerial, &timeout)) { CloseHandle(hSerial); return NULL; }

    PurgeComm(hSerial, PURGE_RXCLEAR | PURGE_TXCLEAR | PURGE_RXABORT | PURGE_TXABORT);
    return hSerial;
}

/* ----------------------------
   Send & Receive
---------------------------- */
int SendData(HANDLE hSerial, uint8_t* data, int length)
{
    DWORD written;
    WriteFile(hSerial, data, length, &written, NULL);
    return (int)written;
}

int ReceiveData(HANDLE hSerial, uint8_t* buffer, int size)
{
    DWORD read;
    ReadFile(hSerial, buffer, size, &read, NULL);
    return (int)read;
}

void PrintHex(uint8_t* data, int len)
{
    int i;
    for (i = 0; i < len; i++) {
        printf("%02X ", data[i]);
    }
    /* 이전 데이터를 덮어쓰기 위해 공백 추가 */
    printf("        \n");
}

/* ----------------------------
   Main
---------------------------- */
int main()
{
    HANDLE serial;
    uint8_t tx[8];
    uint16_t crc;
    uint8_t rx[256];
    int len;
    uint16_t cal_crc, rx_crc, reg_value;

    /* 프로그램 시작 시 화면을 한 번만 완전히 지움 */
    system("cls");

    SetColor(COLOR_CYAN);
    printf("System Booting...\n");
    SetColor(COLOR_DEFAULT);

    serial = OpenSerial("\\\\.\\COM3");

    if (serial == NULL)
    {
        SetColor(COLOR_RED);
        printf("\n[Error] Cannot open COM port.\n");
        SetColor(COLOR_DEFAULT);
        printf("1. Please check if the USB cable is connected.\n");
        printf("2. Verify the COM port number (e.g., COM3) in Device Manager.\n\n");
        system("pause");
        return -1;
    }

    /* 송신 데이터 세팅 (ID:1, Func:3, Addr:0, Count:1) */
    tx[0] = 0x01; tx[1] = 0x03; tx[2] = 0x00; tx[3] = 0x00; tx[4] = 0x00; tx[5] = 0x01;
    crc = ModbusCRC16(tx, 6);
    tx[6] = crc & 0xFF; tx[7] = (crc >> 8) & 0xFF;

    /* 진입 전 화면 클리어 */
    system("cls");

    while (1)
    {
        memset(rx, 0, sizeof(rx));
        SendData(serial, tx, 8);
        Sleep(200);
        len = ReceiveData(serial, rx, sizeof(rx));

        /* 화면을 지우지 않고 커서만 맨 위(0,0)로 이동하여 덮어쓰기 (깜빡임 제거) */
        MoveCursor(0, 0);

        SetColor(COLOR_CYAN);
        printf("==================================================\n");
        printf("            BH-DX103 GAS MONITOR DASHBOARD        \n");
        printf("==================================================\n\n");
        SetColor(COLOR_DEFAULT);

        if (len > 0)
        {
            if (len >= 4)
            {
                cal_crc = ModbusCRC16(rx, len - 2);
                rx_crc = rx[len - 2] | (rx[len - 1] << 8);

                if (cal_crc == rx_crc)
                {
                    if (rx[1] == 0x03 && len == 7)
                    {
                        reg_value = (rx[3] << 8) | rx[4];

                        printf("   [ Status ] : ");
                        SetColor(COLOR_GREEN);
                        printf("Connected (OK)            \n");
                        SetColor(COLOR_DEFAULT);

                        printf("   [ Value  ] : ");
                        SetColor(COLOR_YELLOW);
                        /* %-10d는 숫자가 작아질 때 남는 잔상을 지우기 위한 여백 처리 */
                        printf("%-10d \n\n", reg_value);
                        SetColor(COLOR_DEFAULT);
                    }
                    else
                    {
                        printf("   [ Status ] : ");
                        SetColor(COLOR_RED);
                        printf("Unknown Response Format   \n");
                        printf("   [ Value  ] : --         \n\n");
                        SetColor(COLOR_DEFAULT);
                    }
                }
                else
                {
                    printf("   [ Status ] : ");
                    SetColor(COLOR_RED);
                    printf("CRC Error (Data Corrupted)\n");
                    printf("   [ Value  ] : --         \n\n");
                    SetColor(COLOR_DEFAULT);
                }
            }
            else
            {
                printf("   [ Status ] : ");
                SetColor(COLOR_RED);
                printf("Incomplete Data Received  \n");
                printf("   [ Value  ] : --         \n\n");
                SetColor(COLOR_DEFAULT);
            }
        }
        else
        {
            printf("   [ Status ] : ");
            SetColor(COLOR_RED);
            printf("No Response               \n");
            SetColor(COLOR_DEFAULT);
            printf("                Check Power & A/B wiring  \n");
            printf("   [ Value  ] : --         \n\n");
        }

        SetColor(COLOR_CYAN);
        printf("--------------------------------------------------\n");
        SetColor(COLOR_DEFAULT);

        printf(" * RX Hex Data : ");
        PrintHex(rx, len);
        printf(" * Press 'Ctrl + C' to exit.                      \n");

        SetColor(COLOR_CYAN);
        printf("==================================================\n");
        SetColor(COLOR_DEFAULT);

        Sleep(800);
    }

    CloseHandle(serial);
    return 0;
}
