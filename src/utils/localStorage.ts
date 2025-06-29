import { Housemaid } from '../types/housemaid';

const STORAGE_KEY = 'housemaid_database';
const COUNTER_KEY = 'housemaid_counter';

export const saveHousemaids = (housemaids: Housemaid[]): void => {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(housemaids));
};

export const loadHousemaids = (): Housemaid[] => {
  const data = localStorage.getItem(STORAGE_KEY);
  const housemaids = data ? JSON.parse(data) : [];
  
  // Migrate existing records to include housemaid numbers if they don't have them
  let hasChanges = false;
  const migratedHousemaids = housemaids.map((housemaid: Housemaid) => {
    if (!housemaid.housemaidNumber && housemaid.personalInfo.name.trim()) {
      hasChanges = true;
      return {
        ...housemaid,
        housemaidNumber: generateHousemaidNumber()
      };
    }
    return housemaid;
  });

  // Save migrated data if changes were made
  if (hasChanges) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(migratedHousemaids));
  }

  return migratedHousemaids;
};

export const generateHousemaidNumber = (): string => {
  const currentYear = new Date().getFullYear();
  const counter = getNextCounter();
  return `HM-${currentYear}-${counter.toString().padStart(3, '0')}`;
};

export const generateHousemaidNumberIfEligible = (name: string): string => {
  // Only generate housemaid number if name is provided and not empty
  if (!name || !name.trim()) {
    return '';
  }
  return generateHousemaidNumber();
};

const getNextCounter = (): number => {
  const currentYear = new Date().getFullYear();
  const counterData = localStorage.getItem(COUNTER_KEY);
  
  if (counterData) {
    const { year, count } = JSON.parse(counterData);
    if (year === currentYear) {
      const newCount = count + 1;
      localStorage.setItem(COUNTER_KEY, JSON.stringify({ year: currentYear, count: newCount }));
      return newCount;
    }
  }
  
  // Reset counter for new year or first time
  localStorage.setItem(COUNTER_KEY, JSON.stringify({ year: currentYear, count: 1 }));
  return 1;
};

export const getHousemaidByNumber = (housemaidNumber: string): Housemaid | undefined => {
  const housemaids = loadHousemaids();
  return housemaids.find(h => h.housemaidNumber === housemaidNumber);
};

// Function to check if a housemaid number should be generated
export const shouldGenerateHousemaidNumber = (name: string, currentNumber: string): boolean => {
  // If there's already a number, keep it
  if (currentNumber && currentNumber.trim()) {
    return false;
  }
  
  // Only generate if name is provided and not empty
  return Boolean(name && name.trim().length > 0);
};