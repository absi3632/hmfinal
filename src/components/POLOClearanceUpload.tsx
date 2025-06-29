import React, { useRef, useState } from 'react';
import { Upload, FileText, Eye, Download, Trash2, Shield } from 'lucide-react';

interface POLOClearanceUploadProps {
  clearance: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
    completionDate?: string;
  };
  onUpload: (clearanceData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => void;
  onRemove: () => void;
  onView: () => void;
  completionDate?: string;
  onCompletionDateChange: (date: string) => void;
}

const POLOClearanceUpload: React.FC<POLOClearanceUploadProps> = ({ 
  clearance, 
  onUpload, 
  onRemove, 
  onView,
  completionDate,
  onCompletionDateChange
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isUploading, setIsUploading] = useState(false);

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Validate file type - only PDF and Word documents
    const allowedTypes = [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    ];

    if (!allowedTypes.includes(file.type)) {
      alert('Please upload a PDF or Word document (.pdf, .doc, .docx).');
      return;
    }

    // Validate file size (10MB limit)
    if (file.size > 10 * 1024 * 1024) {
      alert('File size must be less than 10MB.');
      return;
    }

    setIsUploading(true);

    try {
      const reader = new FileReader();
      reader.onload = () => {
        const base64Data = reader.result as string;
        onUpload({
          fileName: file.name,
          fileData: base64Data,
          fileType: file.type,
          uploadDate: new Date().toISOString()
        });
        setIsUploading(false);
      };
      reader.onerror = () => {
        alert('Error reading file. Please try again.');
        setIsUploading(false);
      };
      reader.readAsDataURL(file);
    } catch (error) {
      alert('Error uploading file. Please try again.');
      setIsUploading(false);
    }

    // Reset input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleDownload = () => {
    if (!clearance.fileData || !clearance.fileName) return;

    try {
      const link = document.createElement('a');
      link.href = clearance.fileData;
      link.download = clearance.fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading file. Please try again.');
    }
  };

  const formatFileSize = (base64String: string): string => {
    try {
      const sizeInBytes = (base64String.length * 3) / 4;
      if (sizeInBytes < 1024) return `${sizeInBytes.toFixed(0)} B`;
      if (sizeInBytes < 1024 * 1024) return `${(sizeInBytes / 1024).toFixed(1)} KB`;
      return `${(sizeInBytes / (1024 * 1024)).toFixed(1)} MB`;
    } catch {
      return 'Unknown size';
    }
  };

  const formatDate = (dateString: string): string => {
    return new Date(dateString).toLocaleDateString();
  };

  const getFileTypeIcon = (fileType?: string) => {
    if (fileType === 'application/pdf') {
      return <FileText className="h-6 w-6 text-red-600" />;
    }
    return <FileText className="h-6 w-6 text-blue-600" />;
  };

  if (clearance.fileName && clearance.fileData) {
    return (
      <div className="space-y-4">
        {/* Completion Date Input */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            POLO Clearance Completion Date
          </label>
          <input
            type="date"
            value={completionDate || ''}
            onChange={(e) => onCompletionDateChange(e.target.value)}
            className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
          />
        </div>

        {/* File Display */}
        <div className="border-2 border-dashed border-green-300 rounded-lg p-6 bg-green-50">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="p-2 bg-green-100 rounded-lg">
                {getFileTypeIcon(clearance.fileType)}
              </div>
              <div>
                <p className="font-medium text-gray-900">{clearance.fileName}</p>
                <p className="text-sm text-gray-500">
                  {formatFileSize(clearance.fileData)} • Uploaded {clearance.uploadDate ? formatDate(clearance.uploadDate) : 'Unknown date'}
                </p>
                {completionDate && (
                  <p className="text-sm text-green-600 font-medium">
                    Completed: {formatDate(completionDate)}
                  </p>
                )}
              </div>
            </div>
            <div className="flex items-center space-x-2">
              <button
                type="button"
                onClick={onView}
                className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                title="View POLO Clearance"
              >
                <Eye className="h-4 w-4" />
              </button>
              <button
                type="button"
                onClick={handleDownload}
                className="p-2 text-gray-500 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors"
                title="Download POLO Clearance"
              >
                <Download className="h-4 w-4" />
              </button>
              <button
                type="button"
                onClick={onRemove}
                className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                title="Remove POLO Clearance"
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Completion Date Input */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-2">
          POLO Clearance Completion Date
        </label>
        <input
          type="date"
          value={completionDate || ''}
          onChange={(e) => onCompletionDateChange(e.target.value)}
          className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
        />
      </div>

      {/* File Upload */}
      <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 hover:border-green-400 transition-colors">
        <div className="text-center">
          <div className="mx-auto h-12 w-12 text-gray-400 mb-4">
            {isUploading ? (
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-green-600"></div>
            ) : (
              <div className="flex items-center justify-center">
                <Shield className="h-8 w-8 text-green-600 mr-2" />
                <Upload className="h-8 w-8" />
              </div>
            )}
          </div>
          <div className="mb-4">
            <p className="text-lg font-medium text-gray-900 mb-1">
              {isUploading ? 'Uploading POLO Clearance...' : 'Upload POLO Clearance'}
            </p>
            <p className="text-sm text-gray-500">
              PDF or Word documents up to 10MB
            </p>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.doc,.docx"
            onChange={handleFileSelect}
            disabled={isUploading}
            className="hidden"
          />
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={isUploading}
            className="bg-green-600 text-white px-6 py-2 rounded-lg hover:bg-green-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2 mx-auto"
          >
            <Upload className="h-4 w-4" />
            <span>{isUploading ? 'Uploading...' : 'Choose File'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};

export default POLOClearanceUpload;