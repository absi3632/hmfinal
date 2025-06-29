import React, { useRef, useState } from 'react';
import { Upload, Plane, Eye, Download, Trash2, Ticket } from 'lucide-react';

interface AirTicketUploadProps {
  airTicket: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
    ticketNumber?: string;
    bookingReference?: string;
  };
  onUpload: (ticketData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => void;
  onRemove: () => void;
  onView: () => void;
  ticketNumber?: string;
  bookingReference?: string;
  onTicketNumberChange: (ticketNumber: string) => void;
  onBookingReferenceChange: (bookingReference: string) => void;
}

const AirTicketUpload: React.FC<AirTicketUploadProps> = ({ 
  airTicket, 
  onUpload, 
  onRemove, 
  onView,
  ticketNumber,
  bookingReference,
  onTicketNumberChange,
  onBookingReferenceChange
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isUploading, setIsUploading] = useState(false);

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Validate file type - PDF, images, and Word documents
    const allowedTypes = [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'image/jpeg',
      'image/png',
      'image/jpg',
      'image/webp'
    ];

    if (!allowedTypes.includes(file.type)) {
      alert('Please upload a PDF, Word document, or image file (JPG, PNG, WebP).');
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
    if (!airTicket.fileData || !airTicket.fileName) return;

    try {
      const link = document.createElement('a');
      link.href = airTicket.fileData;
      link.download = airTicket.fileName;
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
      return <Ticket className="h-6 w-6 text-red-600" />;
    }
    if (fileType?.startsWith('image/')) {
      return <Ticket className="h-6 w-6 text-blue-600" />;
    }
    return <Ticket className="h-6 w-6 text-purple-600" />;
  };

  if (airTicket.fileName && airTicket.fileData) {
    return (
      <div className="space-y-4">
        {/* Ticket Details Input */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Ticket Number
            </label>
            <input
              type="text"
              value={ticketNumber || ''}
              onChange={(e) => onTicketNumberChange(e.target.value)}
              className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
              placeholder="e.g., 176-1234567890"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Booking Reference
            </label>
            <input
              type="text"
              value={bookingReference || ''}
              onChange={(e) => onBookingReferenceChange(e.target.value)}
              className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
              placeholder="e.g., ABC123, PNR123456"
            />
          </div>
        </div>

        {/* File Display */}
        <div className="border-2 border-dashed border-purple-300 rounded-lg p-6 bg-purple-50">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="p-2 bg-purple-100 rounded-lg">
                {getFileTypeIcon(airTicket.fileType)}
              </div>
              <div>
                <p className="font-medium text-gray-900">{airTicket.fileName}</p>
                <p className="text-sm text-gray-500">
                  {formatFileSize(airTicket.fileData)} • Uploaded {airTicket.uploadDate ? formatDate(airTicket.uploadDate) : 'Unknown date'}
                </p>
                <div className="flex items-center space-x-4 mt-1">
                  {ticketNumber && (
                    <p className="text-sm text-purple-600 font-medium">
                      Ticket: {ticketNumber}
                    </p>
                  )}
                  {bookingReference && (
                    <p className="text-sm text-purple-600 font-medium">
                      Ref: {bookingReference}
                    </p>
                  )}
                </div>
              </div>
            </div>
            <div className="flex items-center space-x-2">
              <button
                type="button"
                onClick={onView}
                className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                title="View Air Ticket"
              >
                <Eye className="h-4 w-4" />
              </button>
              <button
                type="button"
                onClick={handleDownload}
                className="p-2 text-gray-500 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors"
                title="Download Air Ticket"
              >
                <Download className="h-4 w-4" />
              </button>
              <button
                type="button"
                onClick={onRemove}
                className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                title="Remove Air Ticket"
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
      {/* Ticket Details Input */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Ticket Number
          </label>
          <input
            type="text"
            value={ticketNumber || ''}
            onChange={(e) => onTicketNumberChange(e.target.value)}
            className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
            placeholder="e.g., 176-1234567890"
          />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Booking Reference
          </label>
          <input
            type="text"
            value={bookingReference || ''}
            onChange={(e) => onBookingReferenceChange(e.target.value)}
            className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
            placeholder="e.g., ABC123, PNR123456"
          />
        </div>
      </div>

      {/* File Upload */}
      <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 hover:border-purple-400 transition-colors">
        <div className="text-center">
          <div className="mx-auto h-12 w-12 text-gray-400 mb-4">
            {isUploading ? (
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-purple-600"></div>
            ) : (
              <div className="flex items-center justify-center">
                <Plane className="h-8 w-8 text-purple-600 mr-2" />
                <Upload className="h-8 w-8" />
              </div>
            )}
          </div>
          <div className="mb-4">
            <p className="text-lg font-medium text-gray-900 mb-1">
              {isUploading ? 'Uploading Air Ticket...' : 'Upload Air Ticket'}
            </p>
            <p className="text-sm text-gray-500">
              PDF, Word documents, or images up to 10MB
            </p>
          </div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.doc,.docx,.jpg,.jpeg,.png,.webp"
            onChange={handleFileSelect}
            disabled={isUploading}
            className="hidden"
          />
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={isUploading}
            className="bg-purple-600 text-white px-6 py-2 rounded-lg hover:bg-purple-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2 mx-auto"
          >
            <Upload className="h-4 w-4" />
            <span>{isUploading ? 'Uploading...' : 'Choose File'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};

export default AirTicketUpload;