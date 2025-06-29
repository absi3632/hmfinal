import React, { useRef, useState } from 'react';
import { Upload, File, Eye, Download, Trash2 } from 'lucide-react';

interface CVUploadProps {
  cv: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
  };
  onUpload: (cvData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => void;
  onRemove: () => void;
  onView: () => void;
}

const CVUpload: React.FC<CVUploadProps> = ({ cv, onUpload, onRemove, onView }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isUploading, setIsUploading] = useState(false);

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Validate file type
    const allowedTypes = [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'image/jpeg',
      'image/png',
      'image/jpg'
    ];

    if (!allowedTypes.includes(file.type)) {
      alert('Please upload a PDF, Word document, or image file (JPG, PNG).');
      return;
    }

    // Validate file size (5MB limit)
    if (file.size > 5 * 1024 * 1024) {
      alert('File size must be less than 5MB.');
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
    if (!cv.fileData || !cv.fileName) return;

    try {
      const link = document.createElement('a');
      link.href = cv.fileData;
      link.download = cv.fileName;
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

  if (cv.fileName && cv.fileData) {
    return (
      <div className="border-2 border-dashed border-gray-300 rounded-lg p-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="p-2 bg-blue-100 rounded-lg">
              <File className="h-6 w-6 text-blue-600" />
            </div>
            <div>
              <p className="font-medium text-gray-900">{cv.fileName}</p>
              <p className="text-sm text-gray-500">
                {formatFileSize(cv.fileData)} • Uploaded {cv.uploadDate ? formatDate(cv.uploadDate) : 'Unknown date'}
              </p>
            </div>
          </div>
          <div className="flex items-center space-x-2">
            <button
              type="button"
              onClick={onView}
              className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
              title="View CV"
            >
              <Eye className="h-4 w-4" />
            </button>
            <button
              type="button"
              onClick={handleDownload}
              className="p-2 text-gray-500 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors"
              title="Download CV"
            >
              <Download className="h-4 w-4" />
            </button>
            <button
              type="button"
              onClick={onRemove}
              className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
              title="Remove CV"
            >
              <Trash2 className="h-4 w-4" />
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 hover:border-blue-400 transition-colors">
      <div className="text-center">
        <div className="mx-auto h-12 w-12 text-gray-400 mb-4">
          {isUploading ? (
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
          ) : (
            <Upload className="h-12 w-12" />
          )}
        </div>
        <div className="mb-4">
          <p className="text-lg font-medium text-gray-900 mb-1">
            {isUploading ? 'Uploading CV...' : 'Upload CV'}
          </p>
          <p className="text-sm text-gray-500">
            PDF, Word documents, or images up to 5MB
          </p>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          accept=".pdf,.doc,.docx,.jpg,.jpeg,.png"
          onChange={handleFileSelect}
          disabled={isUploading}
          className="hidden"
        />
        <button
          type="button"
          onClick={() => fileInputRef.current?.click()}
          disabled={isUploading}
          className="bg-blue-600 text-white px-6 py-2 rounded-lg hover:bg-blue-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed"
        >
          {isUploading ? 'Uploading...' : 'Choose File'}
        </button>
      </div>
    </div>
  );
};

export default CVUpload;