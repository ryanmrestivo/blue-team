///////////////////////////////////////////////////////////////////////////////////////////////////
//
//  WinAudit Resources Header
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

#ifndef WINAUDIT_RESOURCES_H_
#define WINAUDIT_RESOURCES_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Application Icon: Defined as 1 so it has the lowest ID
///////////////////////////////////////////////////////////////////////////////////////////////////

#define PXS_IDI_WINAUDIT_FREEWARE                  1

///////////////////////////////////////////////////////////////////////////////////////////////////
// String Identifiers
///////////////////////////////////////////////////////////////////////////////////////////////////

// Main Menu                                             //ENGLISH
const DWORD PXS_IDS_1000_FILE                   = 1000;  // &File
const DWORD PXS_IDS_1001_EDIT                   = 1001;  // &Edit
const DWORD PXS_IDS_1002_VIEW                   = 1002;  // &View
const DWORD PXS_IDS_1003_LANGUAGE               = 1003;  // &Language
const DWORD PXS_IDS_1004_HELP                   = 1004;  // &Help

// File Menu
const DWORD PXS_IDS_1010_AUDIT                  = 1010;  // Audit
const DWORD PXS_IDS_1011_SAVE                   = 1011;  // Save
const DWORD PXS_IDS_1012_SEND_EMAIL             = 1012;  // Send E-Mail
const DWORD PXS_IDS_1013_DATABASE_EXPORT        = 1013;  // Database Export
const DWORD PXS_IDS_1014_DESKTOP_SHORTCUT       = 1014;  // Desktop Shortcut

// Edit Menu
const DWORD PXS_IDS_1020_FIND                   = 1020;  // Find
const DWORD PXS_IDS_1021_FONT                   = 1021;  // Font
const DWORD PXS_IDS_1022_SMALLER_TEXT           = 1022;  // Smaller Text
const DWORD PXS_IDS_1023_BIGGER_TEXT            = 1023;  // Bigger Text

// View Menu
const DWORD PXS_IDS_1030_TOOLBAR                = 1030;  // Toolbar
const DWORD PXS_IDS_1031_STATUSBAR              = 1031;  // Statusbar
const DWORD PXS_IDS_1032_OPTIONS                = 1032;  // Options
const DWORD PXS_IDS_1033_DISK_INFORMATION       = 1033;  // Disk Information
const DWORD PXS_IDS_1034_DISPLAY_INFORMATION    = 1034;  // Display Information
const DWORD PXS_IDS_1035_FIRMWARE_INFORMATION   = 1035;  // Firmware Information
const DWORD PXS_IDS_1036_POLICY_INFORMATION     = 1036;  // Policy Information
const DWORD PXS_IDS_1037_PROCESSOR_INFORMATION  = 1037;  // Processor Information    [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1038_SOFTWARE_INFORMATION   = 1038;  // Software Information

// Help Menu
const DWORD PXS_IDS_1050_USING_WINAUDIT         = 1050;  // Using WinAudit
const DWORD PXS_IDS_1051_ABOUT_WINAUDIT         = 1051;  // About WinAudit
const DWORD PXS_IDS_1052_CREATE_GUID            = 1052;  // Create GUID
const DWORD PXS_IDS_1053_START_LOGGING          = 1053;  // Start Logging
const DWORD PXS_IDS_1054_STOP_LOGGING           = 1054;  // Stop Logging

// Toolbar
const DWORD PXS_IDS_1060_AUDIT_YOUR_COMPUTER    = 1060;  // Audit your computer
const DWORD PXS_IDS_1061_STOP                   = 1061;  // Stop
const DWORD PXS_IDS_1062_STOP_THE_AUDIT         = 1062;  // Stop the audit
const DWORD PXS_IDS_1063_AUDIT_OPTIONS          = 1063;  // Audit Options
const DWORD PXS_IDS_1064_SAVE_TO_FILE           = 1064;  // Save to file
const DWORD PXS_IDS_1065_HELP                   = 1065;  // Help
const DWORD PXS_IDS_1066_VIEW_HELP              = 1066;  // View help

// Database Export Dialog
const DWORD PXS_IDS_1080_EXPORT_AUDIT_TO_DB     = 1080;  // Export Audit to Database         [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1081_DATABASE               = 1081;  // Database
const DWORD PXS_IDS_1082_DATABASE_NAME_COLON    = 1082;  // Database Name:                   [whitespace/lables] NOLINT
const DWORD PXS_IDS_1083_BROWSE                 = 1083;  // Browse
const DWORD PXS_IDS_1084_EXPORT_DATA_OPTIONS    = 1084;  // Export Data Options
const DWORD PXS_IDS_1085_MAXIMUM_ERROS_PERCENT  = 1085;  // Maximum errors [%]:              [whitespace/lables] NOLINT
const DWORD PXS_IDS_1086_MAX_AFFECTED_ROWS      = 1086;  // Max. affected rows:              [whitespace/lables] NOLINT
const DWORD PXS_IDS_1087_CONNECTION_OPTIONS     = 1087;  // Connection Options
const DWORD PXS_IDS_1088_CONNECTION_TIMEOUT_SECS= 1088;  // Connect timeout [secs]:          [whitespace/lables] NOLINT
const DWORD PXS_IDS_1089_QUERY_TIMEOUT_SECS     = 1089;  // Query timeout [secs]:            [whitespace/lables] NOLINT
const DWORD PXS_IDS_1090_ADMINISTRATION         = 1090;  // Administration
const DWORD PXS_IDS_1091_EXPORT                 = 1091;  // Export

// Database Administration Dialog
const DWORD PXS_IDS_1100_CREATE_DATABASE        = 1100;  // Create Database
const DWORD PXS_IDS_1101_CREATE                 = 1101;  // Create
const DWORD PXS_IDS_1102_DATA_MAINTENANCE       = 1102;  // Data Maintenance
const DWORD PXS_IDS_1103_DELETE_OLD_AUDITS      = 1103;  // Delete old audits
const DWORD PXS_IDS_1104_DELETE                 = 1104;  // Delete
const DWORD PXS_IDS_1105_REPORTS                = 1105;  // Reports
const DWORD PXS_IDS_1106_REPORT_NAME            = 1106;  // Report name:                    [whitespace/lables] NOLINT
const DWORD PXS_IDS_1107_MAXIMUM_ROWS           = 1107;  // Maximum rows:                   [whitespace/lables] NOLINT
const DWORD PXS_IDS_1108_SHOW_SQL               = 1108;  // Show SQL
const DWORD PXS_IDS_1109_RUN                    = 1109;  // Run
const DWORD PXS_IDS_1110_ADMINISTRATIVE_TASKS   = 1110;  // Administrative Tasks

// Audit Options Dialog
const DWORD PXS_IDS_1120_SELECT_CATEGORIES      = 1120;  // Select Categories for Auditing   [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1121_SYSTEM_OVERVIEW        = 1121;  // System Overview
const DWORD PXS_IDS_1122_INSTALLED_SOFTWARE     = 1122;  // Installed Software
const DWORD PXS_IDS_1123_OPERATING_SYSTEM       = 1123;  // Operating System
const DWORD PXS_IDS_1124_PERIPHERALS            = 1124;  // Peripherals
const DWORD PXS_IDS_1125_SECURITY               = 1125;  // Security
const DWORD PXS_IDS_1126_GROUPS_AND_USERS       = 1126;  // Groups and Users
const DWORD PXS_IDS_1127_SCHEDULED_TASKS        = 1127;  // Scheduled Tasks
const DWORD PXS_IDS_1128_UPTIME_STATISITCS      = 1128;  // Uptime Statistics
const DWORD PXS_IDS_1129_ERROR_LOGS             = 1129;  // Error Logs
const DWORD PXS_IDS_1130_ENVIRONMENT_VARIABLES  = 1130;  // Environment Variables            [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1131_REGIONAL_SETTINGS      = 1131;  // Regional Settings
const DWORD PXS_IDS_1132_WINDOWS_NETWORK        = 1132;  // Windows Network
const DWORD PXS_IDS_1133_NETWORK_TCPIP          = 1133;  // Network TCP/IP
const DWORD PXS_IDS_1134_HARDWARE_DEVICES       = 1134;  // Hardware Devices
const DWORD PXS_IDS_1135_DISPLAYS               = 1135;  // Displays
const DWORD PXS_IDS_1136_DISPLAY_ADAPTERS       = 1136;  // Display Adapters
const DWORD PXS_IDS_1137_INSTALLED_PRINTERS     = 1137;  // Installed Printers
const DWORD PXS_IDS_1138_BIOS_VERSION           = 1138;  // BIOS Version
const DWORD PXS_IDS_1139_SYSTEM_MANAGEMENT      = 1139;  // System Management
const DWORD PXS_IDS_1140_PROCESSORS             = 1140;  // Processors
const DWORD PXS_IDS_1141_MEMORY                 = 1141;  // Memory
const DWORD PXS_IDS_1142_PHYSICAL_DISKS         = 1142;  // Physical Disks
const DWORD PXS_IDS_1143_DRIVES                 = 1143;  // Drives
const DWORD PXS_IDS_1144_COMMUNICATION_PORTS    = 1144;  // Communication Ports
const DWORD PXS_IDS_1145_STARTUP_PROGRAMMES     = 1145;  // Startup Programs
const DWORD PXS_IDS_1146_SERVICES               = 1146;  // Services
const DWORD PXS_IDS_1147_RUNNING_PROGRAMMES     = 1147;  // Running Programs
const DWORD PXS_IDS_1148_ODBC_INFORMATION       = 1148;  // ODBC Information
const DWORD PXS_IDS_1149_OLE_DB_PROVIDERS       = 1149;  // OLE DB Providers
const DWORD PXS_IDS_1150_SOFTWARE_METERING      = 1150;  // Software Metering
const DWORD PXS_IDS_1151_USER_LOGON_STATISTICS  = 1151;  // User Logon Statistics            [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1152_NONE                   = 1152;  // None
const DWORD PXS_IDS_1153_ALL                    = 1153;  // All
const DWORD PXS_IDS_1154_APPLY                  = 1154;  // Apply

// Categories Panel
const DWORD PXS_IDS_1160_CATEGORIES             = 1160;  // Categories
const DWORD PXS_IDS_1161_ACTIVE_SETUP           = 1161;  // Active Setup
const DWORD PXS_IDS_1162_INSTALLED_PROGRAMS     = 1162;  // Installed Programs
const DWORD PXS_IDS_1163_SOFTWARE_UPDATES       = 1163;  // Software Updates
const DWORD PXS_IDS_1164_KERBEROS_POLICY        = 1164;  // Kerberos Policy
const DWORD PXS_IDS_1165_KERBEROS_TICKETS       = 1165;  // Kerberos Tickets
const DWORD PXS_IDS_1166_NETWORK_TIME_PROTOCOL  = 1166;  // Network Time Protocol            [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1167_OPEN_PORTS             = 1167;  // Open Ports
const DWORD PXS_IDS_1168_PERMISSIONS            = 1168;  // Permissions
const DWORD PXS_IDS_1169_REGISTRY_SECURITY_VALS = 1169;  // Registry Security Values         [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1170_SECURITY_LOG           = 1170;  // Security Log
const DWORD PXS_IDS_1171_SECURITY_SETTINGS      = 1171;  // Security Settings
const DWORD PXS_IDS_1172_SYSTEM_RESTORE         = 1172;  // System Restore
const DWORD PXS_IDS_1173_USER_RIGHTS_ASSIGNMENT = 1173;  // User Rights Assignment           [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1174_USER_PRIVILEGES        = 1174;  // User Privileges
const DWORD PXS_IDS_1175_WINDOWS_FIREWALL       = 1175;  // Windows Firewall
const DWORD PXS_IDS_1176_GROUPS                 = 1176;  // Groups
const DWORD PXS_IDS_1177_GROUP_MEMBERS          = 1177;  // Group Members
const DWORD PXS_IDS_1178_GROUP_POLICY           = 1178;  // Group Policy
const DWORD PXS_IDS_1179_USERS                  = 1179;  // Users
const DWORD PXS_IDS_1180_NETWORK_FILES          = 1180;  // Network Files
const DWORD PXS_IDS_1181_NETWORK_SESSIONS       = 1181;  // Network Sessions
const DWORD PXS_IDS_1182_NETWORK_SHARES         = 1182;  // Network Shares
const DWORD PXS_IDS_1183_BIOS_DETAILS           = 1183;  // BIOS Details
const DWORD PXS_IDS_1184_SYSTEM_INFORMATION     = 1184;  // System Information
const DWORD PXS_IDS_1185_BASE_BOARD             = 1185;  // Base Board
const DWORD PXS_IDS_1186_CHASSIS                = 1186;  // Chassis
const DWORD PXS_IDS_1187_PROCESSOR              = 1187;  // Processor
const DWORD PXS_IDS_1188_MEMORY_CONTROLLER      = 1188;  // Memory Controller
const DWORD PXS_IDS_1189_MEMORY_MODULE          = 1189;  // Memory Module
const DWORD PXS_IDS_1190_CACHE                  = 1190;  // Cache
const DWORD PXS_IDS_1191_PORT_CONNECTOR         = 1191;  // Port Connector
const DWORD PXS_IDS_1192_MEMORY_SLOTS           = 1192;  // System Slots
const DWORD PXS_IDS_1193_MEMORY_ARRAY           = 1193;  // Memory Array
const DWORD PXS_IDS_1194_MEMORY_DEVICE          = 1194;  // Memory Device
const DWORD PXS_IDS_1195_ODBC_DATA_SOURCES      = 1195;  // ODBC Data Sources
const DWORD PXS_IDS_1196_ODBC_DRIVERS           = 1196;  // ODBC Drivers
const DWORD PXS_IDS_1197_NETWORK_ADAPTERS       = 1197;  // Network Adapters
const DWORD PXS_IDS_1198_ROUTING_TABLE          = 1198;  // Routing Table

// Tabs
const DWORD PXS_IDS_1210_AUDIT_REPORT_COMPUTER  = 1210;  // Audit report of your computer    [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1211_DISKS                  = 1211;  // Disks
const DWORD PXS_IDS_1212_DISPLAYS               = 1212;  // Displays
const DWORD PXS_IDS_1213_FIRMWARE               = 1213;  // Firmware
const DWORD PXS_IDS_1214_POLICY                 = 1214;  // Policy
const DWORD PXS_IDS_1215_SOFTWARE               = 1215;  // Software
const DWORD PXS_IDS_1216_LOG                    = 1216;  // Log
const DWORD PXS_IDS_1217_LOGGER_MESSAGES        = 1217;  // Logger messages

// Database Export Messages
const DWORD PXS_IDS_1230_NO_DB_NAME             = 1230;  // A database name was not specified.   [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1231_FULL_PATH_REQUIRED     = 1231;  // A full path is required.             [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1232_USE_DB_NAME_ONLY       = 1232;  // Use the database's name only.        [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1233_DB_NOT_CONNECTED       = 1233;  // Not connected to the database.       [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1234_EXPORTING_RECORDS      = 1234;  // Exporting records
const DWORD PXS_IDS_1235_TOO_MANY_ERRORS        = 1235;  // Too many errors.
const DWORD PXS_IDS_1236_INSERT_RESULT_MESSAGE  = 1236;  // Records inserted: %%1, error count: %%2, time taken: %%3ms   [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1237_TOO_MANY_ROWS_AFFECTED = 1237;  // Too many records will be affected.           [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1238_DATABASE_WAS_CREATED   = 1238;  // The database was successfully created.       [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1239_USE_MDB_OR_ACCDB       = 1239;  // Use either .mdb or .accdb file extension.    [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1240_DELETE_OLD_AUDITS      = 1240;  // This action will delete all but the most recent audit for each computer. Press Yes to continue or No to cancel.  [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1241_NO_OLD_AUDITS          = 1241;  // There is no old audits to delete.            [whitespace/line_length] NOLINT
const DWORD PXS_IDS_1242_DELETED_OLD_AUDITS     = 1242;  // Deleted old audits.
const DWORD PXS_IDS_1243_CONFIRM_EMPTY_DB       = 1243;  // WinAudit requires an empty PostgreSQL database in which it will create tables etc. Use a tool such as pgAdmin to create a database named '%%1'. Confirm this has been done before proceeding.   [whitespace/line_length] NOLINT

// Miscellaneous
const DWORD PXS_IDS_1260_NO                     = 1260;  // No
const DWORD PXS_IDS_1261_YES                    = 1261;  // Yes
const DWORD PXS_IDS_1262_HZ                     = 1262;  // Hz
const DWORD PXS_IDS_1263_MHZ                    = 1263;  // MHz
const DWORD PXS_IDS_1264_BIT                    = 1264;  // bit
const DWORD PXS_IDS_1265_MBS                    = 1265;  // Mbs
const DWORD PXS_IDS_1266_NS                     = 1266;  // ns
const DWORD PXS_IDS_1267_DPI                    = 1267;  // dpi
const DWORD PXS_IDS_1268_COLLAPSE_ALL           = 1268;  // Collapse All
const DWORD PXS_IDS_1269_EXPAND_ALL             = 1269;  // Expand All
const DWORD PXS_IDS_1270_SELECT_ALL             = 1270;  // Collapse All
const DWORD PXS_IDS_1271_DESELECT_ALL           = 1271;  // Deselect All

///////////////////////////////////////////////////////////////////////////////////////////////////
// Icons
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
// Cursors
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
// Bitmaps
///////////////////////////////////////////////////////////////////////////////////////////////////

#define IDB_CLOSE_16                    1000
#define IDB_CLOSE_ON_16                 1001
#define IDB_LOG_FILE_16                 1002
#define IDB_NOT_SAVED_16                1003
#define IDB_SAVE_16                     1004
#define IDB_TREE_NODE_CLOSED            1005
#define IDB_TREE_NODE_OPEN              1006
#define IDB_TREE_LEAF                   1007

// Bitmaps of country flags - ISO 3166-1
#define IDB_FLAG_BE_16                  1008        // Belgium
#define IDB_FLAG_BR_16                  1009        // Brazil
#define IDB_FLAG_CZ_16                  1010        // Czech Republic
#define IDB_FLAG_DA_16                  1011        // Denmark
#define IDB_FLAG_DE_16                  1012        // Germany
#define IDB_FLAG_ES_16                  1013        // Spain
#define IDB_FLAG_FI_16                  1014        // Finland
#define IDB_FLAG_FR_16                  1015        // France
#define IDB_FLAG_GB_16                  1016        // UK
#define IDB_FLAG_GR_16                  1017        // Greece
#define IDB_FLAG_HU_16                  1018        // Hungary
#define IDB_FLAG_KR_16                  1019        // Korea (South)
#define IDB_FLAG_JP_16                  1020        // Japan
#define IDB_FLAG_NL_16                  1021        // Netherlands
#define IDB_FLAG_ID_16                  1022        // Indonesia
#define IDB_FLAG_IL_16                  1023        // Israel
#define IDB_FLAG_IT_16                  1024        // Italy
#define IDB_FLAG_PL_16                  1025        // Poland
#define IDB_FLAG_PT_16                  1026        // Portugal
#define IDB_FLAG_RU_16                  1027        // Russian Federation
#define IDB_FLAG_RS_16                  1028        // Serbia
#define IDB_FLAG_SK_16                  1029        // Slovakia
#define IDB_FLAG_TH_16                  1030        // Thailand
#define IDB_FLAG_TR_16                  1031        // Turkey
#define IDB_FLAG_TW_16                  1032        // Taiwan

///////////////////////////////////////////////////////////////////////////////////////////////////
// Raw Data
///////////////////////////////////////////////////////////////////////////////////////////////////

// Language files - ISO 639-1
#define IDR_STRINGS_WINAUDIT_CS         2001        // Czech
#define IDR_STRINGS_WINAUDIT_DA         2002        // Danish
#define IDR_STRINGS_WINAUDIT_DE         2003        // German
#define IDR_STRINGS_WINAUDIT_EL         2004        // Greek
#define IDR_STRINGS_WINAUDIT_ES         2005        // Spanish
#define IDR_STRINGS_WINAUDIT_EN         2006        // English (UK)
#define IDR_STRINGS_WINAUDIT_FI         2007        // Finnish
#define IDR_STRINGS_WINAUDIT_FR_BE      2008        // French (Belgium)
#define IDR_STRINGS_WINAUDIT_FR_FR      2009        // French (France)
#define IDR_STRINGS_WINAUDIT_HE         2010        // Hebrew
#define IDR_STRINGS_WINAUDIT_HU         2011        // Hungarian
#define IDR_STRINGS_WINAUDIT_ID         2012        // Indonesian
#define IDR_STRINGS_WINAUDIT_IT         2013        // Italian
#define IDR_STRINGS_WINAUDIT_JP         2014        // Japanese
#define IDR_STRINGS_WINAUDIT_KO         2015        // Korean
#define IDR_STRINGS_WINAUDIT_NL         2016        // Dutch
#define IDR_STRINGS_WINAUDIT_PL         2017        // Polish
#define IDR_STRINGS_WINAUDIT_PT_BR      2018        // Portugese (Brazil)
#define IDR_STRINGS_WINAUDIT_PT_PT      2019        // Portuguese (Portugal)
#define IDR_STRINGS_WINAUDIT_RU         2020        // Russian
#define IDR_STRINGS_WINAUDIT_SR         2021        // Serbian
#define IDR_STRINGS_WINAUDIT_SK         2022        // Slovakian
#define IDR_STRINGS_WINAUDIT_TH         2023        // Thai
#define IDR_STRINGS_WINAUDIT_TR         2024        // Turkish
#define IDR_STRINGS_WINAUDIT_ZH_TW      2025        // Chinese Traditional

// Help
#define IDR_WINAUDIT_HELP_RTF           2030

#endif  // WINAUDIT_RESOURCES_H_
