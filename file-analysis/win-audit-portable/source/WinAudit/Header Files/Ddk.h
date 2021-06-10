///////////////////////////////////////////////////////////////////////////////////////////////////
//
// DDK Declarations Header
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2017 PARMAVEX SERVICES
//
// Licensed under the European Union Public Licence (EUPL), Version 1.1 or -
// as soon they will be approved by the European Commission - subsequent
// versions of the EUPL (the "Licence"). You may not use this work except in
// compliance with the Licence. You may obtain a copy of the Licence at:
//
// http://ec.europa.eu/idabc/eupl
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the Licence is distributed on an "AS IS" basis,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the Licence for the specific language governing permissions and
// limitations under the Licence. This source code is free software. It
// must not be sold, leased, rented, sub-licensed or used for any form of
// monetary recompense whatsoever. This notice must not be removed or altered
// from this source distribution.
//
///////////////////////////////////////////////////////////////////////////////////////////////////

#ifndef WINAUDIT_DDK_H_
#define WINAUDIT_DDK_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface

// 2. C System Files
#include <Ntsecapi.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Ntststus.h
///////////////////////////////////////////////////////////////////////////////////////////////////

#ifndef STATUS_SUCCESS
    #define STATUS_SUCCESS              ((NTSTATUS)0x00000000L)  // ntsubauth
#endif

#define STATUS_OBJECT_NAME_NOT_FOUND    ((NTSTATUS)0xC0000034L)

// Will use a name space called PXSDDK and use the required declares from the
// DDK's header files. This avoids compiling with both the SDK and DDK
// simultaneously. A time goes by more of the DDK seems to creep into the SDK.
namespace PXSDDK
{
///////////////////////////////////////////////////////////////////////////////////////////////////
// ntddk.h

#ifndef MAXIMUM_VOLUME_LABEL_LENGTH
    #define MAXIMUM_VOLUME_LABEL_LENGTH  (32 * sizeof ( WCHAR ) )
#endif

typedef short CSHORT;

typedef enum _IO_ALLOCATION_ACTION {
    KeepObject = 1,
    DeallocateObject,
    DeallocateObjectKeepRegisters
} IO_ALLOCATION_ACTION, *PIO_ALLOCATION_ACTION;

typedef
VOID
(*PKDEFERRED_ROUTINE) (
    IN struct _KDPC *Dpc,
    IN PVOID DeferredContext,
    IN PVOID SystemArgument1,
    IN PVOID SystemArgument2
    );

typedef
IO_ALLOCATION_ACTION
(*PDRIVER_CONTROL) (
    IN struct _DEVICE_OBJECT *DeviceObject,
    IN struct _IRP *Irp,
    IN PVOID MapRegisterBase,
    IN PVOID Context
    );

typedef struct _IO_TIMER *PIO_TIMER;

typedef struct _IO_STATUS_BLOCK {
    union {
        NTSTATUS Status;
        PVOID Pointer;
    };

    ULONG_PTR Information;
} IO_STATUS_BLOCK, *PIO_STATUS_BLOCK;

typedef enum _SECTION_INHERIT {
    ViewShare = 1,
    ViewUnmap = 2
} SECTION_INHERIT;

typedef struct _IO_COMPLETION_CONTEXT {
    PVOID Port;
    PVOID Key;
} IO_COMPLETION_CONTEXT, *PIO_COMPLETION_CONTEXT;

typedef struct _VPB {
    CSHORT Type;
    CSHORT Size;
    USHORT Flags;
    USHORT VolumeLabelLength;   // in bytes
    struct _DEVICE_OBJECT *DeviceObject;
    struct _DEVICE_OBJECT *RealDevice;
    ULONG SerialNumber;
    ULONG ReferenceCount;
    WCHAR VolumeLabel[MAXIMUM_VOLUME_LABEL_LENGTH / sizeof(WCHAR)];
} VPB, *PVPB;

typedef struct _KDEVICE_QUEUE {
    CSHORT Type;
    CSHORT Size;
    LIST_ENTRY DeviceListHead;
    KSPIN_LOCK Lock;
    BOOLEAN Busy;
} KDEVICE_QUEUE, *PKDEVICE_QUEUE, *RESTRICTED_POINTER PRKDEVICE_QUEUE;

typedef struct _KDEVICE_QUEUE_ENTRY {
    LIST_ENTRY DeviceListEntry;
    ULONG SortKey;
    BOOLEAN Inserted;
} KDEVICE_QUEUE_ENTRY,
  *PKDEVICE_QUEUE_ENTRY, *RESTRICTED_POINTER PRKDEVICE_QUEUE_ENTRY;

typedef struct _KDPC {
    CSHORT Type;
    UCHAR Number;
    UCHAR Importance;
    LIST_ENTRY DpcListEntry;
    PKDEFERRED_ROUTINE DeferredRoutine;
    PVOID DeferredContext;
    PVOID SystemArgument1;
    PVOID SystemArgument2;
    PULONG_PTR Lock;
} KDPC, *PKDPC, *RESTRICTED_POINTER PRKDPC;

typedef struct _WAIT_CONTEXT_BLOCK {
    KDEVICE_QUEUE_ENTRY WaitQueueEntry;
    PDRIVER_CONTROL DeviceRoutine;
    PVOID DeviceContext;
    ULONG NumberOfMapRegisters;
    PVOID DeviceObject;
    PVOID CurrentIrp;
    PKDPC BufferChainingDpc;
} WAIT_CONTEXT_BLOCK, *PWAIT_CONTEXT_BLOCK;

typedef struct _DISPATCHER_HEADER {
    UCHAR Type;
    UCHAR Absolute;
    UCHAR Size;
    UCHAR Inserted;
    LONG SignalState;
    LIST_ENTRY WaitListHead;
} DISPATCHER_HEADER;

typedef struct _KEVENT {
    DISPATCHER_HEADER Header;
} KEVENT, *PKEVENT, *RESTRICTED_POINTER PRKEVENT;

typedef struct _SECTION_OBJECT_POINTERS {
    PVOID DataSectionObject;
    PVOID SharedCacheMap;
    PVOID ImageSectionObject;
} SECTION_OBJECT_POINTERS;
typedef SECTION_OBJECT_POINTERS *PSECTION_OBJECT_POINTERS;

typedef struct _DEVICE_OBJECT {
    CSHORT Type;
    USHORT Size;
    LONG ReferenceCount;
    struct _DRIVER_OBJECT *DriverObject;
    struct _DEVICE_OBJECT *NextDevice;
    struct _DEVICE_OBJECT *AttachedDevice;
    struct _IRP *CurrentIrp;
    PIO_TIMER Timer;
    ULONG Flags;                                // See above:  DO_...
    ULONG Characteristics;                      // See ntioapi:  FILE_...
    PVPB Vpb;
    PVOID DeviceExtension;
    DEVICE_TYPE DeviceType;
    CCHAR StackSize;
    union {
        LIST_ENTRY ListEntry;
        WAIT_CONTEXT_BLOCK Wcb;
    } Queue;
    ULONG AlignmentRequirement;
    KDEVICE_QUEUE DeviceQueue;
    KDPC Dpc;

    ULONG ActiveThreadCount;
    PSECURITY_DESCRIPTOR SecurityDescriptor;
    KEVENT DeviceLock;

    USHORT SectorSize;
    USHORT Spare1;

    struct _DEVOBJ_EXTENSION  *DeviceObjectExtension;
    PVOID  Reserved;
} DEVICE_OBJECT;
typedef struct _DEVICE_OBJECT *PDEVICE_OBJECT;  // ntndis

typedef struct _FILE_OBJECT {
    CSHORT Type;
    CSHORT Size;
    PDEVICE_OBJECT DeviceObject;
    PVPB Vpb;
    PVOID FsContext;
    PVOID FsContext2;
    PSECTION_OBJECT_POINTERS SectionObjectPointer;
    PVOID PrivateCacheMap;
    NTSTATUS FinalStatus;
    struct _FILE_OBJECT *RelatedFileObject;
    BOOLEAN LockOperation;
    BOOLEAN DeletePending;
    BOOLEAN ReadAccess;
    BOOLEAN WriteAccess;
    BOOLEAN DeleteAccess;
    BOOLEAN SharedRead;
    BOOLEAN SharedWrite;
    BOOLEAN SharedDelete;
    ULONG Flags;
    UNICODE_STRING FileName;
    LARGE_INTEGER CurrentByteOffset;
    ULONG Waiters;
    ULONG Busy;
    PVOID LastLock;
    KEVENT Lock;
    KEVENT Event;
    PIO_COMPLETION_CONTEXT CompletionContext;
} FILE_OBJECT;
typedef struct _FILE_OBJECT *PFILE_OBJECT;

///////////////////////////////////////////////////////////////////////////////////////////////////
// IOCTL/Devices

#define IOCTL_SCSI_MINIPORT_IDENTIFY    ( (0x0000001b << 16) + 0x0501 )

// IDENTIFY_DEVICE_DATA Structure from Windows DDK
#pragma pack( 1 )
typedef struct _IDENTIFY_DEVICE_DATA {
    struct {
    USHORT  Reserved1 : 1;
    USHORT  Retired3 : 1;
    USHORT  ResponseIncomplete : 1;
    USHORT  Retired2 : 3;
    USHORT  FixedDevice : 1;
    USHORT  RemovableMedia : 1;
    USHORT  Retired1 : 7;
    USHORT  DeviceType : 1;
    } GeneralConfiguration;         // word 0
    USHORT  NumCylinders;           // word 1
    USHORT  ReservedWord2;
    USHORT  NumHeads;               // word 3
    USHORT  Retired1[2];
    USHORT  NumSectorsPerTrack;     // word 6
    USHORT  VendorUnique1[3];
    UCHAR   SerialNumber[20];       // word 10-19
    USHORT  Retired2[2];
    USHORT  Obsolete1;
    UCHAR  FirmwareRevision[8];     // word 23-26
    UCHAR  ModelNumber[40];         // word 27-46
    UCHAR  MaximumBlockTransfer;    // word 47
    UCHAR  VendorUnique2;
    USHORT  ReservedWord48;
    struct {
    UCHAR  ReservedByte49;
    UCHAR  DmaSupported : 1;
    UCHAR  LbaSupported : 1;
    UCHAR  IordyDisable : 1;
    UCHAR  IordySupported : 1;
    UCHAR  Reserved1 : 1;
    UCHAR  StandybyTimerSupport : 1;
    UCHAR  Reserved2 : 2;
    USHORT  ReservedWord50;
    } Capabilities;                     // word 49-50
    USHORT  ObsoleteWords51[2];
    USHORT  TranslationFieldsValid:3;   // word 53
    USHORT  Reserved3:13;
    USHORT  NumberOfCurrentCylinders;   // word 54
    USHORT  NumberOfCurrentHeads;       // word 55
    USHORT  CurrentSectorsPerTrack;     // word 56
    ULONG  CurrentSectorCapacity;       // word 57
    UCHAR  CurrentMultiSectorSetting;   // word 58
    UCHAR  MultiSectorSettingValid : 1;
    UCHAR  ReservedByte59 : 7;
    ULONG  UserAddressableSectors;      // word 60-61
    USHORT  ObsoleteWord62;
    USHORT  MultiWordDMASupport : 8;    // word 63
    USHORT  MultiWordDMAActive : 8;
    USHORT  AdvancedPIOModes : 8;
    USHORT  ReservedByte64 : 8;
    USHORT  MinimumMWXferCycleTime;
    USHORT  RecommendedMWXferCycleTime;
    USHORT  MinimumPIOCycleTime;
    USHORT  MinimumPIOCycleTimeIORDY;
    USHORT  ReservedWords69[6];
    USHORT  QueueDepth : 5;
    USHORT  ReservedWord75 : 11;
    USHORT  ReservedWords76[4];
    USHORT  MajorRevision;
    USHORT  MinorRevision;
    struct {
    USHORT  SmartCommands : 1;
    USHORT  SecurityMode : 1;
    USHORT  RemovableMedia : 1;
    USHORT  PowerManagement : 1;
    USHORT  Reserved1 : 1;
    USHORT  WriteCache : 1;
    USHORT  LookAhead : 1;
    USHORT  ReleaseInterrupt : 1;
    USHORT  ServiceInterrupt : 1;
    USHORT  DeviceReset : 1;
    USHORT  HostProtectedArea : 1;
    USHORT  Obsolete1 : 1;
    USHORT  WriteBuffer : 1;
    USHORT  ReadBuffer : 1;
    USHORT  Nop : 1;
    USHORT  Obsolete2 : 1;
    USHORT  DownloadMicrocode : 1;
    USHORT  DmaQueued : 1;
    USHORT  Cfa : 1;
    USHORT  AdvancedPm : 1;
    USHORT  Msn : 1;
    USHORT  PowerUpInStandby : 1;
    USHORT  ManualPowerUp : 1;
    USHORT  Reserved2 : 1;
    USHORT  SetMax : 1;
    USHORT  Acoustics : 1;
    USHORT  BigLba : 1;
    USHORT  Resrved3 : 5;
    } CommandSetSupport;    // word 82-83
    USHORT  ReservedWord84;
    struct {
    USHORT  SmartCommands : 1;
    USHORT  SecurityMode : 1;
    USHORT  RemovableMedia : 1;
    USHORT  PowerManagement : 1;
    USHORT  Reserved1 : 1;
    USHORT  WriteCache : 1;
    USHORT  LookAhead : 1;
    USHORT  ReleaseInterrupt : 1;
    USHORT  ServiceInterrupt : 1;
    USHORT  DeviceReset : 1;
    USHORT  HostProtectedArea : 1;
    USHORT  Obsolete1 : 1;
    USHORT  WriteBuffer : 1;
    USHORT  ReadBuffer : 1;
    USHORT  Nop : 1;
    USHORT  Obsolete2 : 1;
    USHORT  DownloadMicrocode : 1;
    USHORT  DmaQueued : 1;
    USHORT  Cfa : 1;
    USHORT  AdvancedPm : 1;
    USHORT  Msn : 1;
    USHORT  PowerUpInStandby : 1;
    USHORT  ManualPowerUp : 1;
    USHORT  Reserved2 : 1;
    USHORT  SetMax : 1;
    USHORT  Acoustics : 1;
    USHORT  BigLba : 1;
    USHORT  Resrved3 : 5;
    } CommandSetActive;             // word 85-86
    USHORT  ReservedWord87;
    USHORT  UltraDMASupport : 8;    // word 88
    USHORT  UltraDMAActive  : 8;
    USHORT  ReservedWord89[4];
    USHORT  HardwareResetResult;
    USHORT  CurrentAcousticValue : 8;
    USHORT  RecommendedAcousticValue : 8;
    USHORT  ReservedWord95[5];
    ULONG  Max48BitLBA[2];          // word 100-103
    USHORT  ReservedWord104[23];
    USHORT  MsnSupport : 2;
    USHORT  ReservedWord127 : 14;
    USHORT  SecurityStatus;
    USHORT  ReservedWord129[126];
    USHORT  Signature : 8;
    USHORT  CheckSum : 8;
} IDENTIFY_DEVICE_DATA, *PIDENTIFY_DEVICE_DATA;
#pragma pack()  // Reset


// See INQUIRYDATA
typedef struct _INQUIRYDATA {
    UCHAR DeviceType : 5;
    UCHAR DeviceTypeQualifier : 3;
    UCHAR DeviceTypeModifier : 7;
    UCHAR RemovableMedia : 1;
    UCHAR Versions;
    UCHAR ResponseDataFormat : 4;
    UCHAR HiSupport : 1;
    UCHAR NormACA : 1;
    UCHAR ReservedBit : 1;
    UCHAR AERC : 1;
    UCHAR AdditionalLength;
    UCHAR Reserved[2];
    UCHAR SoftReset : 1;
    UCHAR CommandQueue : 1;
    UCHAR Reserved2 : 1;
    UCHAR LinkedCommands : 1;
    UCHAR Synchronous : 1;
    UCHAR Wide16Bit : 1;
    UCHAR Wide32Bit : 1;
    UCHAR RelativeAddressing : 1;
    UCHAR VendorId[8];
    UCHAR ProductId[16];
    UCHAR ProductRevisionLevel[4];
    UCHAR VendorSpecific[20];
    UCHAR Reserved3[40];
} INQUIRYDATA, *PINQUIRYDATA;


#define DIRECT_ACCESS_DEVICE            0x00    // disks
#define SEQUENTIAL_ACCESS_DEVICE        0x01    // tapes
#define PRINTER_DEVICE                  0x02    // printers
#define PROCESSOR_DEVICE                0x03    // scanners, printers, etc
#define WRITE_ONCE_READ_MULTIPLE_DEVICE 0x04    // worms
#define READ_ONLY_DIRECT_ACCESS_DEVICE  0x05    // cdroms
#define SCANNER_DEVICE                  0x06    // scanners
#define OPTICAL_DEVICE                  0x07    // optical disks
#define MEDIUM_CHANGER                  0x08    // jukebox
#define COMMUNICATION_DEVICE            0x09    // network
#define LOGICAL_UNIT_NOT_PRESENT_DEVICE 0x7F

#define DEVICE_QUALIFIER_ACTIVE         0x00
#define DEVICE_QUALIFIER_NOT_ACTIVE     0x01
#define DEVICE_QUALIFIER_NOT_SUPPORTED  0x03

#define DEVICE_CONNECTED 0x00

///////////////////////////////////////////////////////////////////////////////////////////////////
// Structure and Functions for Zwxxx

// The STRING structure is used to pass UNICODE strings
typedef struct _OBJECT_ATTRIBUTES
{
    ULONG           Length;
    HANDLE          RootDirectory;
    PUNICODE_STRING ObjectName;
    ULONG           Attributes;
    PVOID           SecurityDescriptor;
    PVOID           SecurityQualityOfService;
} OBJECT_ATTRIBUTES, *POBJECT_ATTRIBUTES;

///////////////////////////////////////////////////////////////////////////////////////////////////
// SMBIOS

#ifndef ANY_SIZE
    #define ANY_SIZE 1
#endif

// Raw SMBIOS data used  with GetSystemFirmwareTable()
typedef struct RawSMBIOSData
{
    BYTE    Used20CallingMethod;
    BYTE    SMBIOSMajorVersion;
    BYTE    SMBIOSMinorVersion;
    BYTE    DmiRevision;
    DWORD   Length;
    BYTE    SMBIOSTableData[ ANY_SIZE ];
} RAW_SMBIOS_DATA;

}   // namespace PXSDDK

#endif  // WINAUDIT_DDK_H_
