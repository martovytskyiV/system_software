// WMI_App.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

int main(int argc, char **argv)
{
	setlocale(LC_ALL, "");
	HRESULT hres;
	// Крок 1:
	// ініціалізувати COM.
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	
	if (FAILED(hres))
	{
	
		cout << " Не Вдалося Ініціалізувати бібліотеку COM.Помилка коду = 0x"<< hex << hres << endl;
		return 1; // Програма провалилася.
	}
	// Крок 2:
	// Установити загальні COM рівні безпеки
	// Примітка: Вам необхідно вказати - за замовчуванням облікові дані для
	//аутентифікації користувача за допомогою
	// SOLE AUTHENTICATION LISTструктури в pauthlist —
	// параметр Coinitializesecurity
	hres = CoInitializeSecurity(
		NULL,
		-1, // перевірка дійсності COM
		NULL, // Служби перевірки дійсності NULL,
		NULL, // Зарезервований
		RPC_C_AUTHN_LEVEL_DEFAULT, //Перевірка дійсності за замовчуванням
		RPC_C_IMP_LEVEL_IMPERSONATE, //Ідентифікація за замовчуванням
		NULL, // Інформація по аутентифікації
		EOAC_NONE, //Додаткові можливості
		NULL );


	if (FAILED(hres))
	{
		cout << " Не Вдалося безпечно ініціалізувати.Код помилки = 0х"	<< hex << hres << endl;
		CoUninitialize();
		return 1; // Програма закінчена.
	}
	//Крок 3:
	// Одержання первинного локатора для WMI
	IWbemLocator * pLoc = NULL;
	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID*)&pLoc);

	if (FAILED(hres))
	{
		cout << " Не Вдалося створити Iwbemlocator об’єкт."
			<< "Помилка коду = 0x" << hex << hres << endl;
		CoUninitialize();
		return 1; // Завершення програми.
	}
	// Крок 4:
	// Підключення до WMI через Iwbemlocator::Connectserver метод
	IWbemServices * pSvc = NULL;
	//Підключитися до простору імен root\cimv2
	// поточного користувача й одержати покажчик psvc.
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), //Шлях до об’єкта WMI
		NULL, // Ім’я користувача. NULL = поточний користувач.
		NULL, //Пароль користувача. NULL = поточний пароль.
		0,
		NULL,
		0,
		0,
		&pSvc); //покажчик Iwbemservices proxy 
		
	
	if (FAILED(hres))
		{
			cout << " Не Вдалося підключитися. Помилка коду = 0x"
				<< hex << hres << endl; pLoc->Release();
			CoUninitialize();
			return 1; // Завершення програми.
		}
	cout << "Підключений до ROOT\\CIMV2 WMI " << endl;
	// Крок 5:
	// Установити рівень безпеки на проксі
	hres = CoSetProxyBlanket(pSvc, // Покажчик на проксі
		RPC_C_AUTHN_WINNT, // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE, // RPC_C_AUTHZ_xxx
		NULL, // Им’я участника на сервері
		RPC_C_AUTHN_LEVEL_CALL, // RPC_C_AUTHN_LEVEL_xxx
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		NULL, // ідентифікація клієнта
		EOAC_NONE // проксі можливості
	);
	if (FAILED(hres))
	{
		cout << "Не вдалося встановити безпеку на сервері.Помилка коду = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 1; // Завершення програми.
	}
	// Крок 6:
	// Використовувати Iwbemservices покажчик щоб зробити запити WMI --
	// наприклад на одержання імені операційної системи
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY |
		WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);
	if (FAILED(hres))
	{
		cout << "Ім’я операційної системи не отримано."
			<< " помилка коду = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 1; // Програма завершена.
	}
	// Крок 7: -------------------------------------------------
	// Отримати дані з запиту в кроці 6 -------------------
	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;
	while (pEnumerator)
	{
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);
		if (0 == uReturn)
		{
			break;
		}
		VARIANT vtProp;
		VariantInit(&vtProp);
		// Отримати значення властивості Name
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		wcout << " OS Name : " << vtProp.bstrVal << endl;
		VariantClear(&vtProp);
	}
	// Очистка екрана
	// ========
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	pclsObj->Release();
	CoUninitialize();
	return 0; // Програма завершена.
}

