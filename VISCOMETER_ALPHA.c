#define _CRT_SECURE_NO_WARNINGS
#pragma warning(disable: 4819)

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

void convert_to_csv(const char* input_filename, const char* output_filename) {
    FILE* in_file = fopen(input_filename, "r");
    FILE* out_file = NULL;
    int write_header = 0;

    if (in_file == NULL) {
        printf("Error: Cannot find or open file '%s'.\n", input_filename);
        system("pause");
        return;
    }

    // 1. 기존 엑셀 파일이 이미 있는지 확인
    FILE* check_file = fopen(output_filename, "r");
    if (check_file == NULL) {
        // 파일이 없으면 새로 만들어야 하므로 헤더(제목) 작성 필요
        write_header = 1;
    }
    else {
        // 파일이 이미 존재하면 닫기만 함 (헤더 작성 안 함)
        fclose(check_file);
    }

    // 2. 파일을 '이어 쓰기(Append)' 모드("a")로 열기
    out_file = fopen(output_filename, "a");
    if (out_file == NULL) {
        printf("Error: Cannot create or open file '%s'.\n", output_filename);
        fclose(in_file);
        system("pause");
        return;
    }

    // 3. 파일이 처음 만들어질 때만 최상단 헤더 작성
    if (write_header) {
        fprintf(out_file, "Date,Time,Temperature(C),Rotor,Speed(RPM),Viscosity(mPa.s),Percent(%%)\n");
    }

    char line[4096];
    int record_count = 0;
    int error_count = 0;

    while (fgets(line, sizeof(line), in_file)) {
        char* ptr = line;

        while ((ptr = strchr(ptr, '$')) != NULL) {
            int yy, mm, dd, hh, min, ss;
            float temp, viscosity, percent;
            int rotor;
            float speed;

            int match_count = sscanf(ptr + 1, "%d , %d , %d , %d , %d , %d , %f , %d , %f , %f , %f",
                &yy, &mm, &dd, &hh, &min, &ss,
                &temp, &rotor, &speed, &viscosity, &percent);

            if (match_count == 11) {
                fprintf(out_file, "20%02d-%02d-%02d,%02d:%02d:%02d,%.1f,%d,%.1f,%.2f,%.1f\n",
                    yy, mm, dd, hh, min, ss, temp, rotor, speed, viscosity, percent);
                record_count++;
            }
            else {
                error_count++;
            }

            ptr++;
        }
    }

    fclose(in_file);
    fclose(out_file);

    printf("\n=== Conversion Result ===\n");
    printf("- Success: %d records appended\n", record_count);
    printf("- Failed: %d records\n", error_count);
    printf("=========================\n");

    system("pause");
}

int main() {
    const char* input_file = "putty_log.txt";
    const char* output_file = "viscometer_data.csv";

    convert_to_csv(input_file, output_file);
    return 0;
}