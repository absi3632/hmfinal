import React from 'react';

interface StatusBadgeProps {
  status: 'pending' | 'complete' | 'inside' | 'outside' | 'probationary' | 'permanent' | 'resigned' | 'terminated';
  text: string;
}

const StatusBadge: React.FC<StatusBadgeProps> = ({ status, text }) => {
  const getStatusColor = () => {
    switch (status) {
      case 'complete':
      case 'permanent':
        return 'bg-green-100 text-green-800 border-green-200';
      case 'pending':
      case 'probationary':
        return 'bg-orange-100 text-orange-800 border-orange-200';
      case 'inside':
        return 'bg-blue-100 text-blue-800 border-blue-200';
      case 'outside':
        return 'bg-red-100 text-red-800 border-red-200';
      case 'resigned':
        return 'bg-blue-100 text-blue-800 border-blue-200';
      case 'terminated':
        return 'bg-red-100 text-red-800 border-red-200';
      default:
        return 'bg-gray-100 text-gray-800 border-gray-200';
    }
  };

  return (
    <span className={`inline-flex items-center px-3 py-1 rounded-full text-sm font-medium border ${getStatusColor()}`}>
      {text}
    </span>
  );
};

export default StatusBadge;