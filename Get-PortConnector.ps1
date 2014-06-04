# I have no idea where I picked this one up from. -H

﻿function Get-PortConnector {
      $connectiontype = @'
Unknown
Other
Male
Female
Shielded
Unshielded
SCSI (A) High-Density (50 pins)
SCSI (A) Low-Density (50 pins)
SCSI (P) High-Density (68 pins)
SCSI SCA-I (80 pins)
SCSI SCA-II (80 pins)
SCSI Fibre Channel (DB-9, Copper)
SCSI Fibre Channel (Fibre)
SCSI Fibre Channel SCA-II (40 pins)
SCSI Fibre Channel SCA-II (20 pins)
SCSI Fibre Channel BNC
ATA 3-1/2 Inch (40 pins)
ATA 2-1/2 Inch (44 pins)
ATA-2
ATA-3
ATA/66
DB-9
DB-15
DB-25
DB-36
RS-232C
RS-422
RS-423
RS-485
RS-449
V.35
X.21
IEEE-488
AUI
UTP Category 3
UTP Category 4
UTP Category 5
BNC
RJ11
RJ45
Fiber MIC
Apple AUI
Apple GeoPort
PCI
ISA
EISA
VESA
PCMCIA
PCMCIA Type I
PCMCIA Type II
PCMCIA Type III
ZV Port
CardBus
USB
IEEE 1394
HIPPI
HSSDC (6 pins)
GBIC
DIN
Mini-DIN
Micro-DIN
PS/2
Infrared
HP-HIL
Access.bus
NuBus
Centronics
Mini-Centronics
Mini-Centronics Type-14
Mini-Centronics Type-20
Mini-Centronics Type-26
Bus Mouse
ADB
AGP
VME Bus
VME64
Proprietary
Proprietary Processor Card Slot
Proprietary Memory Card Slot
Proprietary I/O Riser Slot
PCI-66MHZ
AGP2X
AGP4X
PC-98
PC-98-Hireso
PC-H98
PC-98Note
PC-98Full
SSA SCSI
Circular
On Board IDE Connector
On Board Floppy Connector
9 Pin Dual Inline
25 Pin Dual Inline
50 Pin Dual Inline
68 Pin Dual Inline
On Board Sound Connector
Mini-Jack
PCI-X
Sbus IEEE 1396-1993 32 Bit
Sbus IEEE 1396-1993 64 Bit
MCA
GIO
XIO
HIO
NGIO
PMC
MTRJ
VF-45
Future I/O
SC
SG
Electrical
Optical
Ribbon
GLM
1x9
Mini SG
LC
HSSC
VHDCI Shielded (68 pins)
InfiniBand
Unknown
Other
Male
Female
Shielded
Unshielded
SCSI (A) High-Density (50 pins)
SCSI (A) Low-Density (50 pins)
SCSI (P) High-Density (68 pins)
SCSI SCA-I (80 pins)
SCSI SCA-II (80 pins)
SCSI Fibre Channel (DB-9, Copper)
SCSI Fibre Channel (Fibre)
SCSI Fibre Channel SCA-II (40 pins)
SCSI Fibre Channel SCA-II (20 pins)
SCSI Fibre Channel BNC
ATA 3-1/2 Inch (40 pins)
ATA 2-1/2 Inch (44 pins)
ATA-2
ATA-3
ATA/66
DB-9
DB-15
DB-25
DB-36
RS-232C
RS-422
RS-423
RS-485
RS-449
V.35
X.21
IEEE-488
AUI
UTP Category 3
UTP Category 4
UTP Category 5
BNC
RJ11
RJ45
Fiber MIC
Apple AUI
Apple GeoPort
PCI
ISA
EISA
VESA
PCMCIA
PCMCIA Type I
PCMCIA Type II
PCMCIA Type III
ZV Port
CardBus
USB
IEEE 1394
HIPPI
HSSDC (6 pins)
GBIC
DIN
Mini-DIN
Micro-DIN
PS/2
Infrared
HP-HIL
Access.bus
NuBus
Centronics
Mini-Centronics
Mini-Centronics Type-14
Mini-Centronics Type-20
Mini-Centronics Type-26
Bus Mouse
ADB
AGP
VME Bus
VME64
Proprietary
Proprietary Processor Card Slot
Proprietary Memory Card Slot
Proprietary I/O Riser Slot
PCI-66MHZ
AGP2X
AGP4X
PC-98
PC-98-Hireso
PC-H98
PC-98Note
PC-98Full
SSA SCSI
Circular
On Board IDE Connector
On Board Floppy Connector
9 Pin Dual Inline
25 Pin Dual Inline
50 Pin Dual Inline
68 Pin Dual Inline
On Board Sound Connector
Mini-Jack
PCI-X
Sbus IEEE 1396-1993 32 Bit
Sbus IEEE 1396-1993 64 Bit
MCA
GIO
XIO
HIO
NGIO
PMC
MTRJ
VF-45
Future I/O
SC
SG
Electrical
Optical
Ribbon
GLM
1x9
Mini SG
LC
HSSC
VHDCI Shielded (68 pins)
InfiniBand
'@.Split(([char]10))

      $porttype = @'
None
Parallel Port XT/AT Compatible
Parallel Port PS/2
Parallel Port ECP
Parallel Port EPP
Parallel Port ECP/EPP
Serial Port XT/AT Compatible
Serial Port 16450 Compatible
Serial Port 16550 Compatible
Serial Port 16550A Compatible
SCSI Port
MIDI Port
Joy Stick Port
Keyboard Port
Mouse Port
SSA SCSI
USB
FireWire (IEEE P1394)
PCMCIA Type II
PCMCIA Type II
PCMCIA Type III
CardBus
Access Bus Port
SCSI II
SCSI Wide
PC-98
PC-98-Hireso
PC-H98
Video Port
Audio Port
Modem Port
Network Port
8251 Compatible
8251 FIFO Compatible
'@.Split(([char]10))

      Get-WmiObject Win32_PortConnector | ForEach-Object {
            $info = @{}
            $info.Name = $_.Tag
            $info.Type = $_.ExternalReferenceDesignator
            $OFS = ", "
            $info.ConnectorType = $_.ConnectorType |
            ForEach-Object { $connectiontype[$_] }
            $info.PortType = $porttype[$_.PortType]

            New-Object PSObject -Property $info
      }
}

Get-PortConnector
