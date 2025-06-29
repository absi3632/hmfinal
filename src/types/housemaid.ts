export interface Housemaid {
  id: string;
  housemaidNumber: string; // Unique housemaid number (e.g., HM-2024-001)
  personalInfo: {
    name: string;
    email: string;
    citizenship: string;
    phone: string; // Combined country code and mobile number
    country: string;
    city: string;
    address: string;
  };
  profilePhoto: {
    fileName?: string;
    fileData?: string; // Base64 encoded image data
    fileType?: string; // MIME type
    uploadDate?: string;
  };
  identity: {
    passportNumber: string;
    passportCountry: string;
    residentId?: string;
  };
  locationStatus: {
    isInsideCountry: boolean;
    exitDate?: string;
    outsideCountryDate?: string; // Date when they went outside the country
  };
  flightInfo: {
    flightDate?: string;
    flightNumber?: string;
    airlineName?: string; // Name of the airline
    destination?: string;
  };
  airTicket: {
    fileName?: string;
    fileData?: string; // Base64 encoded file data
    fileType?: string; // MIME type (PDF, images, etc.)
    uploadDate?: string;
    ticketNumber?: string; // Ticket reference number
    bookingReference?: string; // Booking confirmation code
  };
  employer: {
    name: string;
    mobileNumber: string;
  };
  employment: {
    contractPeriodYears: number;
    startDate: string;
    endDate: string;
    status: 'probationary' | 'permanent' | 'resigned' | 'terminated';
    position?: string;
    salary?: string;
    effectiveDate?: string; // Date when status change became effective (for resigned/terminated)
  };
  recruitmentAgency: {
    name: string;
    licenseNumber?: string;
    contactPerson?: string;
    phoneNumber?: string;
    email?: string;
    address?: string;
  };
  saudiRecruitmentAgency: {
    name: string;
    licenseNumber?: string;
    contactPerson?: string;
    phoneNumber?: string;
    email?: string;
    address?: string;
  };
  complaint: {
    description: string;
    status: 'pending' | 'complete';
    dateReported: string;
    dateResolved?: string;
    resolutionDescription?: string;
  };
  cv: {
    fileName?: string;
    fileData?: string; // Base64 encoded file data
    fileType?: string; // MIME type
    uploadDate?: string;
  };
  poloClearance: {
    fileName?: string;
    fileData?: string; // Base64 encoded file data
    fileType?: string; // MIME type (PDF or Word)
    uploadDate?: string;
    completionDate?: string; // Date when clearance was completed
  };
  createdAt: string;
  updatedAt: string;
}