/*!
 * \author  Hayley Kim
 * \date    20-12-2024
 * \file    Program.cs
 * \brief   Printer program
 * \details Simple program to take zpl template code and replace variable placeholders with variable values, and send as bytes to the printer
 */

/* Includes */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;

/* Namespace */
namespace ZebraPrinter
{
    /*!
     * \brief Class containing main function for sending console input to print function
     */
    internal class Program
    {
        /*!
         * \brief Main function to take inputs from the console and assign them to the appropriate variables in order:
         * 1. Template File name (.prn file)
         * 2. USB Driver for the printer name
         * 3. Variable names in zpl template
         * 4. Variable values to replace variable name placeholders with
         */
        static void Main(string[] args)
        {
            // Arguments taken from console in order
            string templateFilename = args[0];
            string usbDriverName = args[1];
            string variableNames = args[2];
            string variables = args[3];           
            
            // Example values for variables
            //string templateFilename = "C:\\Users\\Tantalus\\Downloads\\testlabel.prn";
            //string usbDriverName = "ZDesigner ZT411-600dpi ZPL";
            //string variableNames = "Product Key^,Barcode^,NID^,PoP^,MD^";
            //string variables = "RT-4230^,0549D7C27E:ECCC618DACA46F7:E010000000000000^,00549D7C2^,ECCC618DACA46F7^,4724^";
            
            // Call to printer function
            Printer.usbprint(variableNames, variables, templateFilename, usbDriverName);

        }

    }

    /*!
     * \brief Class containing functions for parsing zpl code and sending as bytes to printer with Windows Print Spooler service
     */
    public class Printer
    {
        /*!
         * \brief Function to parse zpl code and send to SendStringToPrinter
         */
        public static void usbprint(string variableNames, string variables, string templateFilename, string usbDriverName)
        {
            // Create array from string inputs
            string[] variableNamesArray = ParseStringToArray(variableNames.Trim()); ;
            string[] variablesArray = ParseStringToArray(variables.Trim());

            // Print variables to check
            //Console.WriteLine("Variable Names:");
            //Console.WriteLine(string.Join(", ", variableNamesArray)); // Combines array elements into a single string
            //Console.WriteLine("Variables:");
            //Console.WriteLine(string.Join(", ", variablesArray));

            // Let user know the print service is starting
            Console.WriteLine("\nStarting Print Service");

            try
            {

                // Generate ZPL code by replacing placeholders with variable values
                string zplCode = GenerateZplFromFile(templateFilename, variableNamesArray, variablesArray);

                // Remove newline and carriage return characters from zpl
                string cleanZplCode = RemoveNewline(zplCode);

                // Print the ZPL code to check 
                //Console.WriteLine(zplCode);
                //Console.WriteLine(cleanZplCode);

                // Send cleaned zpl string to the printer
                RawPrinterHelper.SendStringToPrinter(usbDriverName, cleanZplCode);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            // Let user know that print service has ended
            Console.WriteLine("\nPrint Service End");

        }

        /*!
         * \brief Function to load zpl from the file as string
         */
        public static string LoadZplFile(string filePath)
        {
            // Load ZPL file from file path
            if (File.Exists(filePath))
            {
                return File.ReadAllText(filePath);
            }
            else
            {
                throw new FileNotFoundException($"ZPL file not found at: {filePath}");
            }
        }

        /*!
         * \brief Function to go through array of variable names and replace all with variable values from array
         */
        public static string GenerateZplFromFile(string zplFilePath, string[] variableNames, string[] variables)
        {
            // Load the ZPL content from the file by calling LoadZplFile function
            string zplTemplate = LoadZplFile(zplFilePath);

            // Loop through length of array
            for (int i = 0; i < variableNames.Length; i++)
            {
                // Replace the variable name placeholders with the corresponding value from the array
                zplTemplate = zplTemplate.Replace(variableNames[i], variables[i]);
            }

            return zplTemplate;
        }

        /*!
         * \brief Function to remove all newline and carriage return characters
         */
        public static string RemoveNewline(string zplCode)
        {
            // Remove all newline and carriage return characters
            zplCode = zplCode.Replace("\n", "").Replace("\r", "");

            return zplCode;
        }

        /*!
         * \brief Function to split csv string into array
         */
        static string[] ParseStringToArray(string input)
        {
            // Split input string by commas into array
            string[] items = input.Split(',');

            return items;
        }

    }

}

/*!
 * \brief Class containing functions to send raw data to printer
 * \author Microsoft
 */
public class RawPrinterHelper
{
    // Structure and API declarions:
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class DOCINFOA
    {
        [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
        [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
        [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
    }
    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

    [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

    /*
     * \brief   When the function is given a printer name and an unmanaged array of bytes, the function sends those bytes to the print queue.
     * \return  true on success, false on failure.
     */
    public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
    {
        Int32 dwError = 0, dwWritten = 0;
        IntPtr hPrinter = new IntPtr(0);
        DOCINFOA di = new DOCINFOA();
        bool bSuccess = false; // Assume failure unless you specifically succeed.

        di.pDocName = "My C#.NET RAW Document";
        di.pDataType = "RAW";

        // Open the printer.
        if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
        {
            // Start a document.
            if (StartDocPrinter(hPrinter, 1, di))
            {
                // Start a page.
                if (StartPagePrinter(hPrinter))
                {
                    // Write your bytes.
                    bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                    EndPagePrinter(hPrinter);
                }
                EndDocPrinter(hPrinter);
            }
            ClosePrinter(hPrinter);
        }
        // If you did not succeed, GetLastError may give more information
        // about why not.
        if (bSuccess == false)
        {
            dwError = Marshal.GetLastWin32Error();
        }
        return bSuccess;
    }

    public static bool SendFileToPrinter(string szPrinterName, string szFileName)
    {
        // Open the file.
        FileStream fs = new FileStream(szFileName, FileMode.Open);
        // Create a BinaryReader on the file.
        BinaryReader br = new BinaryReader(fs);
        // Dim an array of bytes big enough to hold the file's contents.
        Byte[] bytes = new Byte[fs.Length];
        bool bSuccess = false;
        // Your unmanaged pointer.
        IntPtr pUnmanagedBytes = new IntPtr(0);
        int nLength;

        nLength = Convert.ToInt32(fs.Length);
        // Read the contents of the file into the array.
        bytes = br.ReadBytes(nLength);
        // Allocate some unmanaged memory for those bytes.
        pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
        // Copy the managed byte array into the unmanaged array.
        Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
        // Send the unmanaged bytes to the printer.
        bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
        // Free the unmanaged memory that you allocated earlier.
        Marshal.FreeCoTaskMem(pUnmanagedBytes);
        return bSuccess;
    }
    public static bool SendStringToPrinter(string szPrinterName, string szString)
    {
        try
        {
            IntPtr pBytes;
            Int32 dwCount;
            dwCount = (szString.Length + 1) * Marshal.SystemMaxDBCSCharSize;
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error sending to printer: " + ex.Message);
            return false;
        }
    }

}