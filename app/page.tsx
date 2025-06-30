'use client'

import React, { useState, useEffect } from 'react';

// Load SheetJS (xlsx) library from CDN
// This script will be available globally as 'XLSX'
const loadScript = (src: string, id: string): Promise<void> => { 
  return new Promise((resolve, reject) => {
    if (document.getElementById(id)) {
      return resolve();
    }
    const script = document.createElement('script');
    script.src = src;
    script.id = id;
    script.onload = () => resolve(); // Ensure onload resolves
    script.onerror = (error) => reject(error); // Ensure onerror rejects with error
    document.head.appendChild(script);
  });
};

// Declare XLSX as a global variable so TypeScript knows it exists
declare global {
  interface Window {
    XLSX: unknown; // Using 'any' for simplicity, you could define a more precise type if needed
  }
}

function HomePage() {
  const [excelData, setExcelData] = useState<any[]>([]); // State to store parsed Excel data
  const [fileName, setFileName] = useState<string>(''); // State to store the name of the uploaded file
  const [loading, setLoading] = useState<boolean>(false); // State for loading indicator
  const [error, setError] = useState<string>(''); // State for error messages
  const [headers, setHeaders] = useState<string[]>([]); // State to store table headers

  useEffect(() => {
    // Load the SheetJS library when the component mounts
    loadScript('https://unpkg.com/xlsx/dist/xlsx.full.min.js', 'xlsx-script')
      .then(() => {
        console.log('SheetJS library loaded successfully.');
      })
      .catch(err => {
        setError('Failed to load Excel parsing library. Please try again.');
        console.error('Error loading SheetJS:', err);
      });
  }, []);

  /**
   * Handles the file input change event.
   * Reads the selected Excel file and parses its content.
   * @param {React.ChangeEvent<HTMLInputElement>} event - The file input change event.
   */
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]; // Use optional chaining for files
    if (!file) {
      setError('No file selected.');
      setExcelData([]);
      setFileName('');
      setHeaders([]);
      return;
    }

    // Clear previous errors and set loading state
    setError('');
    setLoading(true);
    setFileName(file.name);

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer); // Type assertion for result
        // Check if XLSX is available globally after loading via CDN
        if (typeof window.XLSX === 'undefined') {
          throw new Error('XLSX library not loaded. Please refresh the page or check your internet connection.');
        }

        // Read the Excel workbook
        const workbook = window.XLSX.read(data, { type: 'array' });

        // Get the first sheet name
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convert the worksheet to a JSON array of objects
        const json: any[][] = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // header: 1 to get array of arrays

        if (json.length === 0) {
          setExcelData([]);
          setHeaders([]);
          setError('The Excel sheet is empty or could not be parsed.');
          return;
        }

        // Assume the first row is the header
        const extractedHeaders: string[] = json[0] as string[];
        const extractedData: any[][] = json.slice(1); // Data starts from the second row

        // Map data rows to objects using extracted headers
        const formattedData = extractedData.map(row => {
          const rowObject: { [key: string]: any } = {}; // Index signature for rowObject
          extractedHeaders.forEach((header, index) => {
            rowObject[header] = row[index];
          });
          return rowObject;
        });

        setHeaders(extractedHeaders);
        setExcelData(formattedData);
      } catch (err: any) { // Catch error as any
        console.error('Error parsing Excel file:', err);
        setError(`Error parsing file: ${err.message}. Please ensure it's a valid Excel file.`);
        setExcelData([]);
        setHeaders([]);
      } finally {
        setLoading(false);
      }
    };

    reader.onerror = (err) => {
      console.error('FileReader error:', err);
      setError('Failed to read file. Please try again.');
      setLoading(false);
      setExcelData([]);
      setHeaders([]);
    };

    reader.readAsArrayBuffer(file); // Read file as ArrayBuffer for XLSX
  };

  return (
    <div className="min-h-screen bg-gray-100 p-4 font-inter">
      <div className="max-w-6xl mx-auto bg-white shadow-lg rounded-xl p-6 md:p-8">
        <h1 className="text-3xl md:text-4xl font-bold text-gray-800 mb-6 text-center">
          Excel Dashboard
        </h1>

        <div className="mb-8 p-4 border border-gray-200 rounded-lg bg-gray-50">
          <label htmlFor="excel-upload" className="block text-lg font-medium text-gray-700 mb-3">
            Upload Excel File
          </label>
          <input
            id="excel-upload"
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="block w-full text-sm text-gray-500
                       file:mr-4 file:py-2 file:px-4
                       file:rounded-full file:border-0
                       file:text-sm file:font-semibold
                       file:bg-blue-50 file:text-blue-700
                       hover:file:bg-blue-100 cursor-pointer"
          />
          {fileName && (
            <p className="mt-2 text-sm text-gray-600">Selected file: <span className="font-semibold">{fileName}</span></p>
          )}
        </div>

        {loading && (
          <div className="flex items-center justify-center py-8">
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-500"></div>
            <p className="ml-4 text-lg text-blue-600">Loading and parsing Excel data...</p>
          </div>
        )}

        {error && (
          <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded-md relative mb-6" role="alert">
            <strong className="font-bold">Error!</strong>
            <span className="block sm:inline ml-2">{error}</span>
          </div>
        )}

        {excelData.length > 0 && (
          <div className="mt-8">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">
              Parsed Data Table
            </h2>
            <div className="overflow-x-auto rounded-lg shadow-md border border-gray-200">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    {headers.map((header, index) => (
                      <th
                        key={index}
                        scope="col"
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider"
                      >
                        {header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {excelData.map((row, rowIndex) => (
                    <tr key={rowIndex} className={rowIndex % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                      {headers.map((header, colIndex) => (
                        <td
                          key={colIndex}
                          className="px-6 py-4 whitespace-nowrap text-sm text-gray-900"
                        >
                          {row[header] !== undefined && row[header] !== null ? String(row[header]) : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {excelData.length === 0 && !loading && !error && fileName && (
          <div className="text-center py-8 text-gray-500">
            <p>No data to display. Please upload an Excel file.</p>
          </div>
        )}
      </div>
    </div>
  );
}

export default HomePage;
