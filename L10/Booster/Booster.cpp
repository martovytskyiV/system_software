#include <windows.h>
#include <stdio.h>
#include "..\L10\PriorityBoosterCommon.h"

int Error(const char* message) {
	printf("%s (error=%d)\n", message, GetLastError());
	return 1;
}


int main(int argc, const char* argv[]) {
	if (argc < 3) {
		printf("Usage: Booster <threadid> <priority>\n");
		return 0;
	}

	HANDLE hDevice = CreateFile(L"\\\\.\\PriorityBooster", GENERIC_WRITE,
		FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
	if (hDevice == INVALID_HANDLE_VALUE)
		return Error("Failed to open device");

	ThreadData data;
	data.ThreadId = atoi(argv[1]); // Первый аргумент командной строки
	data.Priority = atoi(argv[2]); // Второй аргумент командной строк

	DWORD returned;
	BOOL success = DeviceIoControl(hDevice,
		IOCTL_PRIORITY_BOOSTER_SET_PRIORITY, // Код управляющей операции
		&data, sizeof(data), // Входной буфер и длина
		nullptr, 0, // Выходной буфер и длина
		&returned, nullptr);
	if (success)
		printf("Priority change succeeded!\n");
	else
		Error("Priority change failed!");
	CloseHandle(hDevice);

}

