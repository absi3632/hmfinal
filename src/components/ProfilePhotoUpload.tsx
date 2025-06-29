import React, { useRef, useState } from 'react';
import { Upload, Camera, Eye, Download, Trash2 } from 'lucide-react';

interface ProfilePhotoUploadProps {
  photo: {
    fileName?: string;
    fileData?: string;
    fileType?: string;
    uploadDate?: string;
  };
  onUpload: (photoData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => void;
  onRemove: () => void;
  onView: () => void;
}

const ProfilePhotoUpload: React.FC<ProfilePhotoUploadProps> = ({ photo, onUpload, onRemove, onView }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isUploading, setIsUploading] = useState(false);

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Validate file type - only images
    const allowedTypes = [
      'image/jpeg',
      'image/png',
      'image/jpg',
      'image/webp'
    ];

    if (!allowedTypes.includes(file.type)) {
      alert('Please upload an image file (JPG, PNG, WebP).');
      return;
    }

    // Validate file size (2MB limit for photos)
    if (file.size > 2 * 1024 * 1024) {
      alert('Image size must be less than 2MB.');
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
        alert('Error reading image. Please try again.');
        setIsUploading(false);
      };
      reader.readAsDataURL(file);
    } catch (error) {
      alert('Error uploading image. Please try again.');
      setIsUploading(false);
    }

    // Reset input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleDownload = () => {
    if (!photo.fileData || !photo.fileName) return;

    try {
      const link = document.createElement('a');
      link.href = photo.fileData;
      link.download = photo.fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error downloading image. Please try again.');
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

  if (photo.fileName && photo.fileData) {
    return (
      <div className="border-2 border-dashed border-gray-300 rounded-lg p-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-4">
            <div className="relative">
              <img
                src={photo.fileData}
                alt="Profile"
                className="w-16 h-16 rounded-full object-cover border-2 border-gray-200"
              />
            </div>
            <div>
              <p className="font-medium text-gray-900">{photo.fileName}</p>
              <p className="text-sm text-gray-500">
                {formatFileSize(photo.fileData)} • Uploaded {photo.uploadDate ? formatDate(photo.uploadDate) : 'Unknown date'}
              </p>
            </div>
          </div>
          <div className="flex items-center space-x-2">
            <button
              type="button"
              onClick={onView}
              className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
              title="View Photo"
            >
              <Eye className="h-4 w-4" />
            </button>
            <button
              type="button"
              onClick={handleDownload}
              className="p-2 text-gray-500 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors"
              title="Download Photo"
            >
              <Download className="h-4 w-4" />
            </button>
            <button
              type="button"
              onClick={onRemove}
              className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
              title="Remove Photo"
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
        <div className="mx-auto h-20 w-20 text-gray-400 mb-4 flex items-center justify-center">
          {isUploading ? (
            <div className="animate-spin rounded-full h-20 w-20 border-b-2 border-blue-600"></div>
          ) : (
            <div className="w-20 h-20 rounded-full bg-gray-100 flex items-center justify-center">
              <Camera className="h-8 w-8 text-gray-400" />
            </div>
          )}
        </div>
        <div className="mb-4">
          <p className="text-lg font-medium text-gray-900 mb-1">
            {isUploading ? 'Uploading Photo...' : 'Upload Profile Photo'}
          </p>
          <p className="text-sm text-gray-500">
            JPG, PNG, or WebP up to 2MB
          </p>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          accept=".jpg,.jpeg,.png,.webp"
          onChange={handleFileSelect}
          disabled={isUploading}
          className="hidden"
        />
        <button
          type="button"
          onClick={() => fileInputRef.current?.click()}
          disabled={isUploading}
          className="bg-blue-600 text-white px-6 py-2 rounded-lg hover:bg-blue-700 transition-colors font-medium disabled:opacity-50 disabled:cursor-not-allowed flex items-center space-x-2 mx-auto"
        >
          <Upload className="h-4 w-4" />
          <span>{isUploading ? 'Uploading...' : 'Choose Photo'}</span>
        </button>
      </div>
    </div>
  );
};

export default ProfilePhotoUpload;