import React, { useRef, useState } from 'react';
import { Upload, FileSpreadsheet, X, Download, AlertCircle, CheckCircle, Trash2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import { Housemaid } from '../types/housemaid';
import { generateHousemaidNumberIfEligible } from '../utils/localStorage';

interface ExcelImportProps {
  onImport: (housemaids: Housemaid[]) => void;
  onClose: () => void;
}

interface ImportPreview {
  valid: Housemaid[];
  invalid: Array<{
    row: number;
    data: any;
    errors: string[];
  }>;
}

const ExcelImport: React.FC<ExcelImportProps> = ({ onImport, onClose }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [preview, setPreview] = useState<ImportPreview | null>(null);
  const [showPreview, setShowPreview] = useState(false);

  // Expected column headers mapping
  const columnMapping = {
    'Housemaid Number': 'housemaidNumber',
    'Name': 'personalInfo.name',
    'Full Name': 'personalInfo.name',
    'Email': 'personalInfo.email',
    'Email Address': 'personalInfo.email',
    'Citizenship': 'personalInfo.citizenship',
    'Phone': 'personalInfo.phone',
    'Phone Number': 'personalInfo.phone',
    'Country': 'personalInfo.country',
    'City': 'personalInfo.city',
    'Address': 'personalInfo.address',
    'Passport Number': 'identity.passportNumber',
    'Passport Country': 'identity.passportCountry',
    'Resident ID': 'identity.residentId',
    'Location Status': 'locationStatus.isInsideCountry',
    'Inside Country': 'locationStatus.isInsideCountry',
    'Exit Date': 'locationStatus.exitDate',
    'Outside Country Date': 'locationStatus.outsideCountryDate',
    'Flight Date': 'flightInfo.flightDate',
    'Flight Number': 'flightInfo.flightNumber',
    'Airline Name': 'flightInfo.airlineName',
    'Airline': 'flightInfo.airlineName',
    'Destination': 'flightInfo.destination',
    'Employer Name': 'employer.name',
    'Employer Phone': 'employer.mobileNumber',
    'Employer Mobile': 'employer.mobileNumber',
    'Contract Period': 'employment.contractPeriodYears',
    'Contract Years': 'employment.contractPeriodYears',
    'Start Date': 'employment.startDate',
    'End Date': 'employment.endDate',
    'Employment Status': 'employment.status',
    'Position': 'employment.position',
    'Job Title': 'employment.position',
    'Salary': 'employment.salary',
    'Effective Date': 'employment.effectiveDate',
    'Recruitment Agency': 'recruitmentAgency.name',
    'Agency Name': 'recruitmentAgency.name',
    'Agency License': 'recruitmentAgency.licenseNumber',
    'License Number': 'recruitmentAgency.licenseNumber',
    'Agency Contact': 'recruitmentAgency.contactPerson',
    'Contact Person': 'recruitmentAgency.contactPerson',
    'Agency Phone': 'recruitmentAgency.phoneNumber',
    'Agency Email': 'recruitmentAgency.email',
    'Agency Address': 'recruitmentAgency.address',
    'Saudi Agency Name': 'saudiRecruitmentAgency.name',
    'Saudi Recruitment Agency': 'saudiRecruitmentAgency.name',
    'Saudi Agency License': 'saudiRecruitmentAgency.licenseNumber',
    'Saudi License Number': 'saudiRecruitmentAgency.licenseNumber',
    'Saudi Agency Contact': 'saudiRecruitmentAgency.contactPerson',
    'Saudi Contact Person': 'saudiRecruitmentAgency.contactPerson',
    'Saudi Agency Phone': 'saudiRecruitmentAgency.phoneNumber',
    'Saudi Agency Email': 'saudiRecruitmentAgency.email',
    'Saudi Agency Address': 'saudiRecruitmentAgency.address',
    'Complaint Description': 'complaint.description',
    'Complaint Status': 'complaint.status',
    'Date Reported': 'complaint.dateReported',
    'Date Resolved': 'complaint.dateResolved',
    'Resolution Description': 'complaint.resolutionDescription'
  };

  const downloadTemplate = () => {
    const templateData = [
      {
        'Housemaid Number': '', // Leave empty to auto-generate
        'Name': 'Maria Santos',
        'Email': 'maria.santos@email.com',
        'Citizenship': 'Philippines',
        'Phone': '+971 50 123 4567',
        'Country': 'Philippines',
        'City': 'Manila',
        'Address': '123 Main Street, Manila, Philippines',
        'Passport Number': 'P1234567',
        'Passport Country': 'Philippines',
        'Resident ID': 'RES123456',
        'Location Status': 'Inside',
        'Exit Date': '',
        'Outside Country Date': '',
        'Flight Date': '2024-01-15',
        'Flight Number': 'EK123',
        'Airline Name': 'Emirates',
        'Destination': 'Dubai',
        'Employer Name': 'Ahmed Al Mansouri',
        'Employer Phone': '+971 50 987 6543',
        'Contract Period': '2',
        'Start Date': '2024-01-01',
        'End Date': '2026-01-01',
        'Employment Status': 'probationary',
        'Position': 'Housemaid',
        'Salary': 'AED 1,500',
        'Effective Date': '', // Only for resigned/terminated
        'Recruitment Agency': 'ABC Recruitment Services',
        'Agency License': 'LIC-2024-001',
        'Agency Contact': 'Juan Dela Cruz',
        'Agency Phone': '+63 2 123 4567',
        'Agency Email': 'info@abcrecruitment.ph',
        'Agency Address': '456 Recruitment St, Manila, Philippines',
        'Saudi Agency Name': 'Saudi Recruitment Co.',
        'Saudi Agency License': 'SA-LIC-2024-001',
        'Saudi Agency Contact': 'Abdullah Al-Rashid',
        'Saudi Agency Phone': '+966 11 123 4567',
        'Saudi Agency Email': 'info@saudirecruitment.sa',
        'Saudi Agency Address': 'Riyadh, Saudi Arabia',
        'Complaint Description': 'No complaints',
        'Complaint Status': 'pending',
        'Date Reported': '2024-01-01',
        'Date Resolved': '',
        'Resolution Description': ''
      }
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Housemaid Template');
    XLSX.writeFile(wb, 'housemaid_import_template.xlsx');
  };

  const setNestedValue = (obj: any, path: string, value: any) => {
    const keys = path.split('.');
    let current = obj;
    
    for (let i = 0; i < keys.length - 1; i++) {
      if (!(keys[i] in current)) {
        current[keys[i]] = {};
      }
      current = current[keys[i]];
    }
    
    current[keys[keys.length - 1]] = value;
  };

  const validateHousemaid = (data: any, rowIndex: number): { housemaid?: Housemaid; errors: string[] } => {
    const errors: string[] = [];
    
    // Required fields validation
    if (!data.personalInfo?.name?.trim()) {
      errors.push('Name is required');
    }
    if (!data.personalInfo?.phone?.trim()) {
      errors.push('Phone is required');
    }
    if (!data.personalInfo?.address?.trim()) {
      errors.push('Address is required');
    }
    if (!data.identity?.passportNumber?.trim()) {
      errors.push('Passport Number is required');
    }
    if (!data.employer?.name?.trim()) {
      errors.push('Employer Name is required');
    }
    if (!data.employer?.mobileNumber?.trim()) {
      errors.push('Employer Phone is required');
    }
    if (!data.recruitmentAgency?.name?.trim()) {
      errors.push('Recruitment Agency name is required');
    }

    // Email validation
    if (data.personalInfo?.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.personalInfo.email)) {
      errors.push('Invalid email format');
    }

    // Date validation
    const dateFields = [
      { path: 'locationStatus.exitDate', name: 'Exit Date' },
      { path: 'locationStatus.outsideCountryDate', name: 'Outside Country Date' },
      { path: 'flightInfo.flightDate', name: 'Flight Date' },
      { path: 'employment.startDate', name: 'Start Date' },
      { path: 'employment.endDate', name: 'End Date' },
      { path: 'employment.effectiveDate', name: 'Effective Date' },
      { path: 'complaint.dateReported', name: 'Date Reported' },
      { path: 'complaint.dateResolved', name: 'Date Resolved' }
    ];

    dateFields.forEach(field => {
      const value = field.path.split('.').reduce((obj, key) => obj?.[key], data);
      if (value && value !== '' && isNaN(Date.parse(value))) {
        errors.push(`Invalid date format for ${field.name}`);
      }
    });

    // Employment status validation
    if (data.employment?.status && !['probationary', 'permanent', 'resigned', 'terminated'].includes(data.employment.status.toLowerCase())) {
      errors.push('Employment Status must be "probationary", "permanent", "resigned", or "terminated"');
    }

    // Effective date validation for resigned/terminated
    if ((data.employment?.status === 'resigned' || data.employment?.status === 'terminated') && !data.employment?.effectiveDate) {
      errors.push('Effective Date is required for resigned or terminated employees');
    }

    // Complaint status validation
    if (data.complaint?.status && !['pending', 'complete'].includes(data.complaint.status.toLowerCase())) {
      errors.push('Complaint Status must be "pending" or "complete"');
    }

    // Location status validation
    if (data.locationStatus?.isInsideCountry !== undefined) {
      const value = data.locationStatus.isInsideCountry;
      if (typeof value === 'string') {
        const lowerValue = value.toLowerCase();
        if (!['inside', 'outside', 'true', 'false', 'yes', 'no'].includes(lowerValue)) {
          errors.push('Location Status must be "Inside", "Outside", "Yes", "No", "True", or "False"');
        }
      }
    }

    if (errors.length > 0) {
      return { errors };
    }

    // Generate housemaid number if not provided and name is available
    const housemaidNumber = data.housemaidNumber?.trim() || 
      (data.personalInfo?.name?.trim() ? generateHousemaidNumberIfEligible(data.personalInfo.name.trim()) : '');

    // Create housemaid object
    const now = new Date().toISOString();
    const housemaid: Housemaid = {
      id: `import_${Date.now()}_${rowIndex}`,
      housemaidNumber: housemaidNumber,
      personalInfo: {
        name: data.personalInfo?.name?.trim() || '',
        email: data.personalInfo?.email?.trim() || '',
        citizenship: data.personalInfo?.citizenship?.trim() || '',
        phone: data.personalInfo?.phone?.trim() || '',
        country: data.personalInfo?.country?.trim() || '',
        city: data.personalInfo?.city?.trim() || '',
        address: data.personalInfo?.address?.trim() || ''
      },
      profilePhoto: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined
      },
      identity: {
        passportNumber: data.identity?.passportNumber?.trim() || '',
        passportCountry: data.identity?.passportCountry?.trim() || '',
        residentId: data.identity?.residentId?.trim() || ''
      },
      locationStatus: {
        isInsideCountry: parseLocationStatus(data.locationStatus?.isInsideCountry),
        exitDate: data.locationStatus?.exitDate ? formatDate(data.locationStatus.exitDate) : undefined,
        outsideCountryDate: data.locationStatus?.outsideCountryDate ? formatDate(data.locationStatus.outsideCountryDate) : undefined
      },
      flightInfo: {
        flightDate: data.flightInfo?.flightDate ? formatDate(data.flightInfo.flightDate) : undefined,
        flightNumber: data.flightInfo?.flightNumber?.trim() || '',
        airlineName: data.flightInfo?.airlineName?.trim() || '',
        destination: data.flightInfo?.destination?.trim() || ''
      },
      airTicket: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined,
        ticketNumber: '',
        bookingReference: ''
      },
      employer: {
        name: data.employer?.name?.trim() || '',
        mobileNumber: data.employer?.mobileNumber?.trim() || ''
      },
      employment: {
        contractPeriodYears: parseInt(data.employment?.contractPeriodYears) || 2,
        startDate: data.employment?.startDate ? formatDate(data.employment.startDate) : '',
        endDate: data.employment?.endDate ? formatDate(data.employment.endDate) : '',
        status: (data.employment?.status?.toLowerCase() === 'permanent' ? 'permanent' : 
                 data.employment?.status?.toLowerCase() === 'resigned' ? 'resigned' :
                 data.employment?.status?.toLowerCase() === 'terminated' ? 'terminated' : 'probationary') as 'probationary' | 'permanent' | 'resigned' | 'terminated',
        position: data.employment?.position?.trim() || '',
        salary: data.employment?.salary?.trim() || '',
        effectiveDate: data.employment?.effectiveDate ? formatDate(data.employment.effectiveDate) : undefined
      },
      recruitmentAgency: {
        name: data.recruitmentAgency?.name?.trim() || '',
        licenseNumber: data.recruitmentAgency?.licenseNumber?.trim() || '',
        contactPerson: data.recruitmentAgency?.contactPerson?.trim() || '',
        phoneNumber: data.recruitmentAgency?.phoneNumber?.trim() || '',
        email: data.recruitmentAgency?.email?.trim() || '',
        address: data.recruitmentAgency?.address?.trim() || ''
      },
      saudiRecruitmentAgency: {
        name: data.saudiRecruitmentAgency?.name?.trim() || '',
        licenseNumber: data.saudiRecruitmentAgency?.licenseNumber?.trim() || '',
        contactPerson: data.saudiRecruitmentAgency?.contactPerson?.trim() || '',
        phoneNumber: data.saudiRecruitmentAgency?.phoneNumber?.trim() || '',
        email: data.saudiRecruitmentAgency?.email?.trim() || '',
        address: data.saudiRecruitmentAgency?.address?.trim() || ''
      },
      complaint: {
        description: data.complaint?.description?.trim() || '',
        status: (data.complaint?.status?.toLowerCase() === 'complete' ? 'complete' : 'pending') as 'pending' | 'complete',
        dateReported: data.complaint?.dateReported ? formatDate(data.complaint.dateReported) : new Date().toISOString().split('T')[0],
        dateResolved: data.complaint?.dateResolved ? formatDate(data.complaint.dateResolved) : undefined,
        resolutionDescription: data.complaint?.resolutionDescription?.trim() || ''
      },
      cv: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined
      },
      poloClearance: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined,
        completionDate: undefined
      },
      createdAt: now,
      updatedAt: now
    };

    return { housemaid, errors: [] };
  };

  const parseLocationStatus = (value: any): boolean => {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') {
      const lowerValue = value.toLowerCase();
      return ['inside', 'true', 'yes'].includes(lowerValue);
    }
    return true; // default to inside
  };

  const formatDate = (value: any): string => {
    if (!value) return '';
    const date = new Date(value);
    if (isNaN(date.getTime())) return '';
    return date.toISOString().split('T')[0];
  };

  const processExcelFile = async (file: File) => {
    setIsProcessing(true);
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (jsonData.length < 2) {
        throw new Error('Excel file must contain at least a header row and one data row');
      }

      const headers = jsonData[0] as string[];
      const dataRows = jsonData.slice(1);

      const valid: Housemaid[] = [];
      const invalid: Array<{ row: number; data: any; errors: string[] }> = [];

      dataRows.forEach((row: any, index) => {
        if (row.every((cell: any) => !cell || cell === '')) return; // Skip empty rows

        const rowData: any = {};
        
        // Map Excel columns to our data structure
        headers.forEach((header, colIndex) => {
          const mappedPath = columnMapping[header as keyof typeof columnMapping];
          if (mappedPath && row[colIndex] !== undefined && row[colIndex] !== '') {
            setNestedValue(rowData, mappedPath, row[colIndex]);
          }
        });

        const validation = validateHousemaid(rowData, index);
        
        if (validation.housemaid) {
          valid.push(validation.housemaid);
        } else {
          invalid.push({
            row: index + 2, // +2 because we start from row 1 and skip header
            data: rowData,
            errors: validation.errors
          });
        }
      });

      setPreview({ valid, invalid });
      setShowPreview(true);
    } catch (error) {
      alert(`Error processing Excel file: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Validate file type
    const allowedTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
      'text/csv' // .csv
    ];

    if (!allowedTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls|csv)$/i)) {
      alert('Please upload an Excel file (.xlsx, .xls) or CSV file.');
      return;
    }

    // Validate file size (10MB limit)
    if (file.size > 10 * 1024 * 1024) {
      alert('File size must be less than 10MB.');
      return;
    }

    await processExcelFile(file);

    // Reset input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleImport = () => {
    if (preview?.valid) {
      onImport(preview.valid);
      onClose();
    }
  };

  const handleClearPreview = () => {
    setPreview(null);
    setShowPreview(false);
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg max-w-6xl w-full max-h-[90vh] overflow-hidden">
        <div className="flex items-center justify-between p-6 border-b">
          <div>
            <h3 className="text-xl font-semibold text-gray-900">Import from Excel</h3>
            <p className="text-sm text-gray-600 mt-1">Upload an Excel file to bulk import housemaid records</p>
          </div>
          <button
            onClick={onClose}
            className="p-2 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-lg transition-colors"
            title="Close"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        <div className="p-6 overflow-y-auto max-h-[calc(90vh-120px)]">
          {!showPreview ? (
            <div className="space-y-6">
              {/* Template Download */}
              <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
                <div className="flex items-start space-x-3">
                  <FileSpreadsheet className="h-6 w-6 text-blue-600 mt-0.5" />
                  <div className="flex-1">
                    <h4 className="font-medium text-blue-900 mb-1">Download Template</h4>
                    <p className="text-sm text-blue-700 mb-3">
                      Download our Excel template with sample data and proper column headers to ensure successful import.
                    </p>
                    <button
                      onClick={downloadTemplate}
                      className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors font-medium flex items-center space-x-2"
                    >
                      <Download className="h-4 w-4" />
                      <span>Download Template</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Important Notice */}
              <div className="bg-yellow-50 p-4 rounded-lg border border-yellow-200">
                <div className="flex items-start space-x-3">
                  <AlertCircle className="h-6 w-6 text-yellow-600 mt-0.5" />
                  <div className="flex-1">
                    <h4 className="font-medium text-yellow-900 mb-1">Important: Required Fields</h4>
                    <p className="text-sm text-yellow-700">
                      The following fields are required for successful import: Name, Phone, Address, Passport Number, 
                      Employer Name, Employer Phone, and Recruitment Agency Name. A unique housemaid number will be 
                      automatically generated for records with a valid name. For resigned or terminated employees, 
                      an Effective Date is also required.
                    </p>
                  </div>
                </div>
              </div>

              {/* File Upload */}
              <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 hover:border-blue-400 transition-colors">
                <div className="text-center">
                  <div className="mx-auto h-16 w-16 text-gray-400 mb-4">
                    {isProcessing ? (
                      <div className="animate-spin rounded-full h-16 w-16 border-b-2 border-blue-600"></div>
                    ) : (
                      <FileSpreadsheet className="h-16 w-16" />
                    )}
                  </div>
                  <div className="mb-4">
                    <p className="text-lg font-medium text-gray-900 mb-1">
                      {isProcessing ? 'Processing Excel File...' : 'Upload Excel File'}
                    </p>
                    <p className="text-sm text-gray-500">
                      Excel (.xlsx, .xls) or CSV files up to 10MB
                    </p>
                  </div>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleFileSelect}
                    disabled={isProcessing}
                    className="hidden"
                  />
                  <button
                    type="button"
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isProcessing}
                    className="bg-blue-600 text-white px-6 py-3 rounded-lg hover:bg-blue-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2 mx-auto"
                  >
                    <Upload className="h-5 w-5" />
                    <span>{isProcessing ? 'Processing...' : 'Choose File'}</span>
                  </button>
                </div>
              </div>

              {/* Column Mapping Guide */}
              <div className="bg-gray-50 p-4 rounded-lg">
                <h4 className="font-medium text-gray-900 mb-3">Supported Column Headers</h4>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2 text-sm">
                  {Object.keys(columnMapping).map((header) => (
                    <div key={header} className="bg-white px-3 py-1 rounded border">
                      {header}
                    </div>
                  ))}
                </div>
                <p className="text-xs text-gray-600 mt-3">
                  Column headers are case-sensitive. Use the exact headers shown above or download the template.
                  <br />
                  <strong>Note:</strong> If "Housemaid Number" is not provided, it will be automatically generated only if the "Name" field is filled.
                  <br />
                  <strong>Employment Status:</strong> Accepted values are "probationary", "permanent", "resigned", or "terminated".
                  <br />
                  <strong>Effective Date:</strong> Required for "resigned" or "terminated" employment status.
                </p>
              </div>
            </div>
          ) : (
            <div className="space-y-6">
              {/* Import Summary */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="bg-green-50 p-4 rounded-lg border border-green-200">
                  <div className="flex items-center">
                    <CheckCircle className="h-6 w-6 text-green-600 mr-2" />
                    <div>
                      <p className="font-medium text-green-900">Valid Records</p>
                      <p className="text-2xl font-bold text-green-900">{preview?.valid.length || 0}</p>
                    </div>
                  </div>
                </div>
                <div className="bg-red-50 p-4 rounded-lg border border-red-200">
                  <div className="flex items-center">
                    <AlertCircle className="h-6 w-6 text-red-600 mr-2" />
                    <div>
                      <p className="font-medium text-red-900">Invalid Records</p>
                      <p className="text-2xl font-bold text-red-900">{preview?.invalid.length || 0}</p>
                    </div>
                  </div>
                </div>
                <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
                  <div className="flex items-center">
                    <FileSpreadsheet className="h-6 w-6 text-blue-600 mr-2" />
                    <div>
                      <p className="font-medium text-blue-900">Total Processed</p>
                      <p className="text-2xl font-bold text-blue-900">{(preview?.valid.length || 0) + (preview?.invalid.length || 0)}</p>
                    </div>
                  </div>
                </div>
              </div>

              {/* Valid Records Preview */}
              {preview && preview.valid.length > 0 && (
                <div>
                  <h4 className="font-medium text-gray-900 mb-3 flex items-center">
                    <CheckCircle className="h-5 w-5 text-green-600 mr-2" />
                    Valid Records ({preview.valid.length})
                  </h4>
                  <div className="bg-white border rounded-lg overflow-hidden">
                    <div className="max-h-60 overflow-y-auto">
                      <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gray-50">
                          <tr>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Housemaid #</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Name</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Email</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Phone</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Passport</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Employer</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">PH Agency</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">SA Agency</th>
                            <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                          {preview.valid.slice(0, 10).map((housemaid, index) => (
                            <tr key={index} className="hover:bg-gray-50">
                              <td className="px-4 py-2 text-sm font-medium text-blue-900">{housemaid.housemaidNumber}</td>
                              <td className="px-4 py-2 text-sm text-gray-900">{housemaid.personalInfo.name}</td>
                              <td className="px-4 py-2 text-sm text-gray-600">{housemaid.personalInfo.email || '-'}</td>
                              <td className="px-4 py-2 text-sm text-gray-600">{housemaid.personalInfo.phone}</td>
                              <td className="px-4 py-2 text-sm text-gray-600">{housemaid.identity.passportNumber}</td>
                              <td className="px-4 py-2 text-sm text-gray-600">{housemaid.employer.name}</td>
                              <td className="px-4 py-2 text-sm text-yellow-700">{housemaid.recruitmentAgency.name}</td>
                              <td className="px-4 py-2 text-sm text-green-700">{housemaid.saudiRecruitmentAgency.name || '-'}</td>
                              <td className="px-4 py-2 text-sm text-gray-600 capitalize">{housemaid.employment.status}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {preview.valid.length > 10 && (
                      <div className="px-4 py-2 bg-gray-50 text-sm text-gray-600 text-center">
                        ... and {preview.valid.length - 10} more records
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* Invalid Records */}
              {preview && preview.invalid.length > 0 && (
                <div>
                  <h4 className="font-medium text-gray-900 mb-3 flex items-center">
                    <AlertCircle className="h-5 w-5 text-red-600 mr-2" />
                    Invalid Records ({preview.invalid.length})
                  </h4>
                  <div className="bg-white border rounded-lg overflow-hidden">
                    <div className="max-h-60 overflow-y-auto">
                      <div className="space-y-3 p-4">
                        {preview.invalid.map((item, index) => (
                          <div key={index} className="bg-red-50 p-3 rounded border border-red-200">
                            <p className="font-medium text-red-900 mb-1">Row {item.row}</p>
                            <div className="text-sm text-red-700">
                              <p className="mb-1">
                                <span className="font-medium">Name:</span> {item.data.personalInfo?.name || 'Not provided'}
                              </p>
                              <p className="mb-2">
                                <span className="font-medium">Errors:</span>
                              </p>
                              <ul className="list-disc list-inside space-y-1">
                                {item.errors.map((error, errorIndex) => (
                                  <li key={errorIndex}>{error}</li>
                                ))}
                              </ul>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Action Buttons */}
              <div className="flex justify-between items-center pt-4 border-t">
                <button
                  onClick={handleClearPreview}
                  className="px-4 py-2 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors font-medium flex items-center space-x-2"
                >
                  <Trash2 className="h-4 w-4" />
                  <span>Clear Preview</span>
                </button>
                <div className="flex space-x-3">
                  <button
                    onClick={onClose}
                    className="px-6 py-2 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors font-medium"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleImport}
                    disabled={(preview?.valid.length || 0) === 0}
                    className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2"
                  >
                    <Upload className="h-4 w-4" />
                    <span>Import {preview?.valid.length || 0} Records</span>
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default ExcelImport;