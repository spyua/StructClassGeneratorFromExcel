# Produce struct class from excel file.

Support the struct class generator solution. The struct class can use on tcp communcation message definition. It's easy to create the message class speedly.


## Excel Message Format

![image](https://user-images.githubusercontent.com/20264622/107941514-c5ef4e80-6fc4-11eb-8ccb-5e68f10017ac.png)

## Data Type Format
 - IL,DW = int
 - R,F = float
 - I,W = short
 - c,n = byte[]

## Usage
 - Define your message column on Excel
 - Set App config
  - Excel Path:Read excel path
  - OutFileName: Cs file name
  - Export Path: Cs target path
  
  - StarSheet:The read of star excel page
  - StarReadRow:The read of star row
  - StarReadColumn:The read of star column
  
## Result
```
using System;
using System.Runtime.InteropServices;
namespace MsgStruct
{
    public class Msg_Demo
    {
        [Serializable]
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public class Msg_MsgA
        {
            [MarshalAs(UnmanagedType.I2)]
            public short MessageLength;
            [MarshalAs(UnmanagedType.I2)]
            public short MessageId;
            [MarshalAs(UnmanagedType.I4)]
            public int Date;
            [MarshalAs(UnmanagedType.I4)]
            public int Time;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 24)]
            public byte[] ID;
            [MarshalAs(UnmanagedType.I4)]
            public int Mode;
            [MarshalAs(UnmanagedType.R4)]
            public float Length;

        }
        [Serializable]
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public class Msg_MsgB
        {
            [MarshalAs(UnmanagedType.I2)]
            public short MessageLength;
            [MarshalAs(UnmanagedType.I2)]
            public short MessageId;
            [MarshalAs(UnmanagedType.I4)]
            public int Date;
            [MarshalAs(UnmanagedType.I4)]
            public int Time;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare1;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare2;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare3;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare4;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare5;

        }
        [Serializable]
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public class Msg_MsgC
        {
            [MarshalAs(UnmanagedType.I2)]
            public short MessageLength;
            [MarshalAs(UnmanagedType.I2)]
            public short MessageId;
            [MarshalAs(UnmanagedType.I4)]
            public int Date;
            [MarshalAs(UnmanagedType.I4)]
            public int Time;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 24)]
            public byte[] ID;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare1;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare2;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare3;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare4;
            [MarshalAs(UnmanagedType.I4)]
            public int Spare5;

        }

    }
}
```
 
