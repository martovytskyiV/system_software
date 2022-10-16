#include <ntifs.h>
#include <ntddk.h>
#include "PriorityBoosterCommon.h"

extern "C" 

// Прототипи
void PriorityBoosterUnload(_In_ PDRIVER_OBJECT DriverObject);
NTSTATUS PriorityBoosterCreateClose(_In_ PDEVICE_OBJECT DeviceObject, 	_In_ PIRP Irp);
NTSTATUS PriorityBoosterDeviceControl(PDEVICE_OBJECT, PIRP Irp);



UNICODE_STRING devName = RTL_CONSTANT_STRING(L"\\Device\\PriorityBooster");
// RtlInitUnicodeString(&devName, L»\\Device\\ThreadBoost»);

// DriverEntry
extern "C" NTSTATUS
DriverEntry(_In_ PDRIVER_OBJECT DriverObject, _In_ PUNICODE_STRING RegistryPath)
{

	DriverObject->MajorFunction[IRP_MJ_CREATE] = PriorityBoosterCreateClose;
	DriverObject->MajorFunction[IRP_MJ_CLOSE] = PriorityBoosterCreateClose;
	DriverObject->MajorFunction[IRP_MJ_DEVICE_CONTROL] = PriorityBoosterDeviceControl;
	DriverObject->DriverUnload = PriorityBoosterUnload;

	PDEVICE_OBJECT DeviceObject;
	NTSTATUS status = IoCreateDevice(
		DriverObject, // наш объект драйвера,
		0, // дополнительные байты не нужны,
		&devName, // имя устройства,
		FILE_DEVICE_UNKNOWN, // тип устройства,
		0, // флаги характеристик,
		FALSE, // не монопольное использование,
		&DeviceObject // полученный указатель
	);
	if (!NT_SUCCESS(status)) {
		KdPrint(("Failed to create device object (0x%08X)\n", status));
		return status;
	}

	UNICODE_STRING symLink = RTL_CONSTANT_STRING(L"\\??\\PriorityBooster");
	status = IoCreateSymbolicLink(&symLink, &devName);
	if (!NT_SUCCESS(status)) {
		KdPrint(("Failed to create symbolic link (0x%08X)\n", status));
		IoDeleteDevice(DeviceObject);
		return status;
	}


	UNREFERENCED_PARAMETER(RegistryPath);
	return STATUS_SUCCESS;
}






void PriorityBoosterUnload(_In_ PDRIVER_OBJECT DriverObject) {
	UNICODE_STRING symLink = RTL_CONSTANT_STRING(L"\\??\\PriorityBooster");
	// Удалить символическую ссылку
	IoDeleteSymbolicLink(&symLink);
	// Удалить объект устройства
	IoDeleteDevice(DriverObject->DeviceObject);

}



_Use_decl_annotations_
NTSTATUS PriorityBoosterCreateClose(PDEVICE_OBJECT DeviceObject, PIRP Irp) {
	UNREFERENCED_PARAMETER(DeviceObject);
	Irp->IoStatus.Status = STATUS_SUCCESS;
	Irp->IoStatus.Information = 0;
	IoCompleteRequest(Irp, IO_NO_INCREMENT);
	return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS PriorityBoosterDeviceControl(PDEVICE_OBJECT, PIRP Irp) {
	// Получить IO_STACK_LOCATION
	auto stack = IoGetCurrentIrpStackLocation(Irp); // IO_STACK_LOCATION*
	auto status = STATUS_SUCCESS;
	switch (stack->Parameters.DeviceIoControl.IoControlCode) {
	case IOCTL_PRIORITY_BOOSTER_SET_PRIORITY:{
		// Выполнить основную работу
		if (stack->Parameters.DeviceIoControl.InputBufferLength < sizeof(ThreadData)) {
			status = STATUS_BUFFER_TOO_SMALL;
			break;
		}
		auto data = (ThreadData*)stack->Parameters.DeviceIoControl.Type3InputBuffer;
		if (data == nullptr) {
			status = STATUS_INVALID_PARAMETER;
			break;
		}
		if (data->Priority < 1 || data->Priority > 31) {
			status = STATUS_INVALID_PARAMETER;
			break;
		}

		PETHREAD Thread;
		status = PsLookupThreadByThreadId(ULongToHandle(data->ThreadId), &Thread);
		if (!NT_SUCCESS(status))
			break;

		KeSetPriorityThread((PKTHREAD)Thread, data->Priority);
		ObDereferenceObject(Thread);
		KdPrint(("Thread Priority change for %d to %d succeeded!\n",
			data->ThreadId, data->Priority));
		break;
	}
	default:
		status = STATUS_INVALID_DEVICE_REQUEST;
		break;
	}


	Irp->IoStatus.Status = status;
	Irp->IoStatus.Information = 0;
	IoCompleteRequest(Irp, IO_NO_INCREMENT);
	return status;

}