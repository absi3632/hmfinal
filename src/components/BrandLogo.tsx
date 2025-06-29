import React from 'react';
import { Users } from 'lucide-react';
import { BrandSettings } from '../types/brand';

interface BrandLogoProps {
  brandSettings: BrandSettings;
  size?: 'small' | 'medium' | 'large';
  showText?: boolean;
  className?: string;
}

const BrandLogo: React.FC<BrandLogoProps> = ({ 
  brandSettings, 
  size = 'medium', 
  showText = true,
  className = ''
}) => {
  const getSizeClasses = () => {
    switch (size) {
      case 'small':
        return {
          container: 'w-8 h-8',
          icon: 'h-4 w-4',
          text: 'text-sm',
          subtext: 'text-xs'
        };
      case 'large':
        return {
          container: 'w-20 h-20',
          icon: 'h-10 w-10',
          text: 'text-2xl',
          subtext: 'text-sm'
        };
      default: // medium
        return {
          container: 'w-10 h-10',
          icon: 'h-6 w-6',
          text: 'text-xl',
          subtext: 'text-xs'
        };
    }
  };

  const sizeClasses = getSizeClasses();

  const logoStyle = {
    background: brandSettings.logoFileData 
      ? 'transparent' 
      : `linear-gradient(to right, ${brandSettings.primaryColor || '#2563eb'}, ${brandSettings.secondaryColor || '#7c3aed'})`
  };

  return (
    <div className={`flex items-center space-x-3 ${className}`}>
      <div 
        className={`${sizeClasses.container} rounded-lg flex items-center justify-center overflow-hidden`}
        style={logoStyle}
      >
        {brandSettings.logoFileData ? (
          <img
            src={brandSettings.logoFileData}
            alt={brandSettings.companyName || 'Company Logo'}
            className="w-full h-full object-contain"
          />
        ) : (
          <Users className={`${sizeClasses.icon} text-white`} />
        )}
      </div>
      
      {showText && (
        <div>
          <h1 className={`${sizeClasses.text} font-bold text-gray-900`}>
            {brandSettings.companyName || 'Housemaid Management'}
          </h1>
          {brandSettings.tagline && (
            <p className={`${sizeClasses.subtext} text-gray-500`}>
              {brandSettings.tagline}
            </p>
          )}
        </div>
      )}
    </div>
  );
};

export default BrandLogo;