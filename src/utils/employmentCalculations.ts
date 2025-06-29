export interface EmploymentCalculation {
  daysWorked: number;
  shouldBePermanent: boolean;
  daysUntilPermanent: number;
  isEligibleForPermanent: boolean;
}

export const calculateEmploymentStatus = (startDate: string): EmploymentCalculation => {
  if (!startDate) {
    return {
      daysWorked: 0,
      shouldBePermanent: false,
      daysUntilPermanent: 90,
      isEligibleForPermanent: false
    };
  }

  const start = new Date(startDate);
  const today = new Date();
  
  // Calculate days worked
  const diffTime = today.getTime() - start.getTime();
  const daysWorked = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  
  // Check if should be permanent (90 days or more)
  const shouldBePermanent = daysWorked >= 90;
  const daysUntilPermanent = Math.max(0, 90 - daysWorked);
  const isEligibleForPermanent = daysWorked >= 90;

  return {
    daysWorked: Math.max(0, daysWorked), // Don't show negative days
    shouldBePermanent,
    daysUntilPermanent,
    isEligibleForPermanent
  };
};

export const getEmploymentStatusColor = (status: 'probationary' | 'permanent' | 'resigned' | 'terminated', daysWorked: number) => {
  if (status === 'permanent') {
    return 'bg-green-100 text-green-800 border-green-200';
  }
  
  if (status === 'resigned') {
    return 'bg-blue-100 text-blue-800 border-blue-200';
  }
  
  if (status === 'terminated') {
    return 'bg-red-100 text-red-800 border-red-200';
  }
  
  if (daysWorked >= 90) {
    return 'bg-blue-100 text-blue-800 border-blue-200'; // Ready for permanent
  }
  
  if (daysWorked >= 60) {
    return 'bg-yellow-100 text-yellow-800 border-yellow-200'; // Close to permanent
  }
  
  return 'bg-orange-100 text-orange-800 border-orange-200'; // Still probationary
};

export const getEmploymentStatusText = (status: 'probationary' | 'permanent' | 'resigned' | 'terminated', daysWorked: number) => {
  if (status === 'permanent') {
    return 'Permanent Employee';
  }
  
  if (status === 'resigned') {
    return 'Resigned';
  }
  
  if (status === 'terminated') {
    return 'Terminated';
  }
  
  if (daysWorked >= 90) {
    return 'Eligible for Permanent';
  }
  
  return 'Probationary';
};

export const formatDaysWorked = (days: number): string => {
  if (days === 0) return '0 days';
  if (days === 1) return '1 day';
  
  if (days >= 365) {
    const years = Math.floor(days / 365);
    const remainingDays = days % 365;
    if (remainingDays === 0) {
      return `${years} ${years === 1 ? 'year' : 'years'}`;
    }
    return `${years} ${years === 1 ? 'year' : 'years'}, ${remainingDays} ${remainingDays === 1 ? 'day' : 'days'}`;
  }
  
  if (days >= 30) {
    const months = Math.floor(days / 30);
    const remainingDays = days % 30;
    if (remainingDays === 0) {
      return `${months} ${months === 1 ? 'month' : 'months'}`;
    }
    return `${months} ${months === 1 ? 'month' : 'months'}, ${remainingDays} ${remainingDays === 1 ? 'day' : 'days'}`;
  }
  
  return `${days} days`;
};