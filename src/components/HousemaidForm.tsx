import React, { useState, useEffect } from 'react';
import { 
  Save, 
  X, 
  User, 
  Mail, 
  Phone, 
  MapPin, 
  Building, 
  Calendar, 
  FileText, 
  Plane, 
  Shield,
  AlertTriangle,
  Camera,
  Ticket,
  Users
} from 'lucide-react';
import { Housemaid } from '../types/housemaid';
import { countries, passportCountries } from '../data/countries';
import { generateHousemaidNumberIfEligible, shouldGenerateHousemaidNumber } from '../utils/localStorage';
import { hasPermission, getCurrentUser } from '../utils/auth';
import CountrySelect from './CountrySelect';
import ProfilePhotoUpload from './ProfilePhotoUpload';
import ProfilePhotoViewer from './ProfilePhotoViewer';
import CVUpload from './CVUpload';
import CVViewer from './CVViewer';
import POLOClearanceUpload from './POLOClearanceUpload';
import POLOClearanceViewer from './POLOClearanceViewer';
import AirTicketUpload from './AirTicketUpload';
import AirTicketViewer from './AirTicketViewer';

interface HousemaidFormProps {
  housemaid?: Housemaid;
  onSave: (housemaid: Housemaid) => void;
  onCancel: () => void;
}

const HousemaidForm: React.FC<HousemaidFormProps> = ({ housemaid, onSave, onCancel }) => {
  const currentUser = getCurrentUser();
  const isReadOnly = !hasPermission(currentUser?.role || 'viewer', 'manager');
  
  const [formData, setFormData] = useState<Partial<Housemaid>>({
    personalInfo: {
      name: '',
      email: '',
      citizenship: '',
      phone: '',
      country: '',
      city: '',
      address: ''
    },
    profilePhoto: {
      fileName: undefined,
      fileData: undefined,
      fileType: undefined,
      uploadDate: undefined
    },
    identity: {
      passportNumber: '',
      passportCountry: '',
      residentId: ''
    },
    locationStatus: {
      isInsideCountry: true,
      exitDate: undefined,
      outsideCountryDate: undefined
    },
    flightInfo: {
      flightDate: undefined,
      flightNumber: '',
      airlineName: '',
      destination: ''
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
      name: '',
      mobileNumber: ''
    },
    employment: {
      contractPeriodYears: 2,
      startDate: '',
      endDate: '',
      status: 'probationary',
      position: '',
      salary: '',
      effectiveDate: undefined
    },
    recruitmentAgency: {
      name: '',
      licenseNumber: '',
      contactPerson: '',
      phoneNumber: '',
      email: '',
      address: ''
    },
    saudiRecruitmentAgency: {
      name: '',
      licenseNumber: '',
      contactPerson: '',
      phoneNumber: '',
      email: '',
      address: ''
    },
    complaint: {
      description: '',
      status: 'pending',
      dateReported: new Date().toISOString().split('T')[0],
      dateResolved: undefined,
      resolutionDescription: ''
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
    housemaidNumber: ''
  });

  const [showPhotoViewer, setShowPhotoViewer] = useState(false);
  const [showCVViewer, setShowCVViewer] = useState(false);
  const [showPOLOViewer, setShowPOLOViewer] = useState(false);
  const [showAirTicketViewer, setShowAirTicketViewer] = useState(false);

  useEffect(() => {
    if (housemaid) {
      setFormData(housemaid);
    }
  }, [housemaid]);

  const handleInputChange = (section: string, field: string, value: any) => {
    if (isReadOnly) return;
    
    setFormData(prev => ({
      ...prev,
      [section]: {
        ...prev[section as keyof typeof prev],
        [field]: value
      }
    }));

    // Auto-generate housemaid number when name is entered (only for new records)
    if (section === 'personalInfo' && field === 'name' && !housemaid) {
      const currentNumber = (formData as any).housemaidNumber || '';
      if (shouldGenerateHousemaidNumber(value, currentNumber)) {
        setFormData(prev => ({
          ...prev,
          housemaidNumber: generateHousemaidNumberIfEligible(value)
        }));
      }
    }

    // Auto-calculate end date when start date or contract period changes
    if (section === 'employment' && (field === 'startDate' || field === 'contractPeriodYears')) {
      const startDate = field === 'startDate' ? value : formData.employment?.startDate;
      const contractYears = field === 'contractPeriodYears' ? value : formData.employment?.contractPeriodYears;
      
      if (startDate && contractYears) {
        const start = new Date(startDate);
        const end = new Date(start);
        end.setFullYear(start.getFullYear() + parseInt(contractYears));
        
        setFormData(prev => ({
          ...prev,
          employment: {
            ...prev.employment!,
            endDate: end.toISOString().split('T')[0]
          }
        }));
      }
    }
  };

  const handleDirectChange = (field: string, value: any) => {
    if (isReadOnly) return;
    
    setFormData(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const handlePhotoUpload = (photoData: { fileName: string; fileData: string; fileType: string; uploadDate: string }) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      profilePhoto: photoData
    }));
  };

  const handlePhotoRemove = () => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      profilePhoto: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined
      }
    }));
  };

  const handleCVUpload = (cvData: { fileName: string; fileData: string; fileType: string; uploadDate: string }) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      cv: cvData
    }));
  };

  const handleCVRemove = () => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      cv: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined
      }
    }));
  };

  const handlePOLOUpload = (clearanceData: { fileName: string; fileData: string; fileType: string; uploadDate: string }) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      poloClearance: {
        ...prev.poloClearance!,
        ...clearanceData
      }
    }));
  };

  const handlePOLORemove = () => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      poloClearance: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined,
        completionDate: prev.poloClearance?.completionDate
      }
    }));
  };

  const handlePOLOCompletionDateChange = (date: string) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      poloClearance: {
        ...prev.poloClearance!,
        completionDate: date
      }
    }));
  };

  const handleAirTicketUpload = (ticketData: { fileName: string; fileData: string; fileType: string; uploadDate: string }) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      airTicket: {
        ...prev.airTicket!,
        ...ticketData
      }
    }));
  };

  const handleAirTicketRemove = () => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      airTicket: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined,
        ticketNumber: prev.airTicket?.ticketNumber || '',
        bookingReference: prev.airTicket?.bookingReference || ''
      }
    }));
  };

  const handleAirTicketNumberChange = (ticketNumber: string) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      airTicket: {
        ...prev.airTicket!,
        ticketNumber
      }
    }));
  };

  const handleAirTicketReferenceChange = (bookingReference: string) => {
    if (isReadOnly) return;
    setFormData(prev => ({
      ...prev,
      airTicket: {
        ...prev.airTicket!,
        bookingReference
      }
    }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (isReadOnly) return;

    const now = new Date().toISOString();
    const housemaidData: Housemaid = {
      id: housemaid?.id || `housemaid-${Date.now()}`,
      housemaidNumber: formData.housemaidNumber || '',
      personalInfo: formData.personalInfo!,
      profilePhoto: formData.profilePhoto!,
      identity: formData.identity!,
      locationStatus: formData.locationStatus!,
      flightInfo: formData.flightInfo!,
      airTicket: formData.airTicket!,
      employer: formData.employer!,
      employment: formData.employment!,
      recruitmentAgency: formData.recruitmentAgency!,
      saudiRecruitmentAgency: formData.saudiRecruitmentAgency!,
      complaint: formData.complaint!,
      cv: formData.cv!,
      poloClearance: formData.poloClearance!,
      createdAt: housemaid?.createdAt || now,
      updatedAt: now
    };

    onSave(housemaidData);
  };

  return (
    <div className="max-w-6xl mx-auto">
      {/* Header */}
      <div className="bg-gradient-to-r from-blue-600 to-purple-600 rounded-2xl p-8 text-white mb-8">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-4xl font-bold mb-2">
              {housemaid ? 'Edit Housemaid Record' : 'Add New Housemaid'}
            </h1>
            <p className="text-blue-100 text-lg">
              {isReadOnly ? 'View detailed information' : 'Complete all required information to create a comprehensive record'}
            </p>
          </div>
          {formData.profilePhoto?.fileData && (
            <div className="w-24 h-24 rounded-full overflow-hidden border-4 border-white shadow-lg">
              <img
                src={formData.profilePhoto.fileData}
                alt="Profile"
                className="w-full h-full object-cover"
              />
            </div>
          )}
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-8">
        {/* Personal Information */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-blue-100 rounded-xl mr-4">
              <User className="h-6 w-6 text-blue-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Personal Information</h2>
              <p className="text-gray-600">Basic personal details and contact information</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Housemaid Number
              </label>
              <input
                type="text"
                value={formData.housemaidNumber || ''}
                onChange={(e) => handleDirectChange('housemaidNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="Auto-generated when name is entered"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Full Name <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={formData.personalInfo?.name || ''}
                onChange={(e) => handleInputChange('personalInfo', 'name', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                required
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Email Address
              </label>
              <input
                type="email"
                value={formData.personalInfo?.email || ''}
                onChange={(e) => handleInputChange('personalInfo', 'email', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Citizenship/Nationality
              </label>
              <input
                type="text"
                value={formData.personalInfo?.citizenship || ''}
                onChange={(e) => handleInputChange('personalInfo', 'citizenship', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Phone Number <span className="text-red-500">*</span>
              </label>
              <input
                type="tel"
                value={formData.personalInfo?.phone || ''}
                onChange={(e) => handleInputChange('personalInfo', 'phone', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="+971 50 123 4567"
                required
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Country of Residence
              </label>
              <input
                type="text"
                value={formData.personalInfo?.country || ''}
                onChange={(e) => handleInputChange('personalInfo', 'country', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                City
              </label>
              <input
                type="text"
                value={formData.personalInfo?.city || ''}
                onChange={(e) => handleInputChange('personalInfo', 'city', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                readOnly={isReadOnly}
              />
            </div>

            <div className="md:col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Address <span className="text-red-500">*</span>
              </label>
              <textarea
                value={formData.personalInfo?.address || ''}
                onChange={(e) => handleInputChange('personalInfo', 'address', e.target.value)}
                rows={3}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                required
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Profile Photo */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-green-100 rounded-xl mr-4">
              <Camera className="h-6 w-6 text-green-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Profile Photo</h2>
              <p className="text-gray-600">Upload a professional profile photo</p>
            </div>
          </div>

          <ProfilePhotoUpload
            photo={formData.profilePhoto!}
            onUpload={handlePhotoUpload}
            onRemove={handlePhotoRemove}
            onView={() => setShowPhotoViewer(true)}
          />
        </div>

        {/* Identity Information */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-purple-100 rounded-xl mr-4">
              <Shield className="h-6 w-6 text-purple-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Identity Information</h2>
              <p className="text-gray-600">Passport and identification details</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Passport Number <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={formData.identity?.passportNumber || ''}
                onChange={(e) => handleInputChange('identity', 'passportNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
                required
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Passport Issuing Country
              </label>
              <select
                value={formData.identity?.passportCountry || ''}
                onChange={(e) => handleInputChange('identity', 'passportCountry', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
                disabled={isReadOnly}
              >
                <option value="">Select Country</option>
                {passportCountries.map(country => (
                  <option key={country} value={country}>{country}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Resident ID
              </label>
              <input
                type="text"
                value={formData.identity?.residentId || ''}
                onChange={(e) => handleInputChange('identity', 'residentId', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Location Status */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-red-100 rounded-xl mr-4">
              <MapPin className="h-6 w-6 text-red-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Location Status</h2>
              <p className="text-gray-600">Current location and travel information</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Current Status
              </label>
              <select
                value={formData.locationStatus?.isInsideCountry ? 'inside' : 'outside'}
                onChange={(e) => handleInputChange('locationStatus', 'isInsideCountry', e.target.value === 'inside')}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-red-500"
                disabled={isReadOnly}
              >
                <option value="inside">Inside Country</option>
                <option value="outside">Outside Country</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Exit Date
              </label>
              <input
                type="date"
                value={formData.locationStatus?.exitDate || ''}
                onChange={(e) => handleInputChange('locationStatus', 'exitDate', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-red-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Outside Country Date
              </label>
              <input
                type="date"
                value={formData.locationStatus?.outsideCountryDate || ''}
                onChange={(e) => handleInputChange('locationStatus', 'outsideCountryDate', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-red-500"
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Flight Information */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-indigo-100 rounded-xl mr-4">
              <Plane className="h-6 w-6 text-indigo-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Flight Information</h2>
              <p className="text-gray-600">Travel and flight details</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Flight Date
              </label>
              <input
                type="date"
                value={formData.flightInfo?.flightDate || ''}
                onChange={(e) => handleInputChange('flightInfo', 'flightDate', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Flight Number
              </label>
              <input
                type="text"
                value={formData.flightInfo?.flightNumber || ''}
                onChange={(e) => handleInputChange('flightInfo', 'flightNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                placeholder="EK123"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Airline Name
              </label>
              <input
                type="text"
                value={formData.flightInfo?.airlineName || ''}
                onChange={(e) => handleInputChange('flightInfo', 'airlineName', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                placeholder="Emirates"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Destination
              </label>
              <input
                type="text"
                value={formData.flightInfo?.destination || ''}
                onChange={(e) => handleInputChange('flightInfo', 'destination', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                placeholder="Dubai"
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Air Ticket */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-purple-100 rounded-xl mr-4">
              <Ticket className="h-6 w-6 text-purple-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Air Ticket</h2>
              <p className="text-gray-600">Upload air ticket document and details</p>
            </div>
          </div>

          <AirTicketUpload
            airTicket={formData.airTicket!}
            onUpload={handleAirTicketUpload}
            onRemove={handleAirTicketRemove}
            onView={() => setShowAirTicketViewer(true)}
            ticketNumber={formData.airTicket?.ticketNumber}
            bookingReference={formData.airTicket?.bookingReference}
            onTicketNumberChange={handleAirTicketNumberChange}
            onBookingReferenceChange={handleAirTicketReferenceChange}
          />
        </div>

        {/* Employer Information */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-yellow-100 rounded-xl mr-4">
              <Building className="h-6 w-6 text-yellow-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Employer Information</h2>
              <p className="text-gray-600">Current employer details</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Employer Name <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={formData.employer?.name || ''}
                onChange={(e) => handleInputChange('employer', 'name', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                required
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Employer Mobile Number <span className="text-red-500">*</span>
              </label>
              <input
                type="tel"
                value={formData.employer?.mobileNumber || ''}
                onChange={(e) => handleInputChange('employer', 'mobileNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                placeholder="+971 50 123 4567"
                required
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Employment Information */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-green-100 rounded-xl mr-4">
              <Calendar className="h-6 w-6 text-green-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Employment Information</h2>
              <p className="text-gray-600">Contract and employment details</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Position
              </label>
              <input
                type="text"
                value={formData.employment?.position || ''}
                onChange={(e) => handleInputChange('employment', 'position', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                placeholder="Housemaid"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Employment Status
              </label>
              <select
                value={formData.employment?.status || 'probationary'}
                onChange={(e) => handleInputChange('employment', 'status', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                disabled={isReadOnly}
              >
                <option value="probationary">Probationary</option>
                <option value="permanent">Permanent</option>
                <option value="resigned">Resigned</option>
                <option value="terminated">Terminated</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Contract Period (Years)
              </label>
              <select
                value={formData.employment?.contractPeriodYears || 2}
                onChange={(e) => handleInputChange('employment', 'contractPeriodYears', parseInt(e.target.value))}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                disabled={isReadOnly}
              >
                <option value={1}>1 Year</option>
                <option value={2}>2 Years</option>
                <option value={3}>3 Years</option>
                <option value={4}>4 Years</option>
                <option value={5}>5 Years</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Start Date
              </label>
              <input
                type="date"
                value={formData.employment?.startDate || ''}
                onChange={(e) => handleInputChange('employment', 'startDate', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                End Date
              </label>
              <input
                type="date"
                value={formData.employment?.endDate || ''}
                onChange={(e) => handleInputChange('employment', 'endDate', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Salary
              </label>
              <input
                type="text"
                value={formData.employment?.salary || ''}
                onChange={(e) => handleInputChange('employment', 'salary', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                placeholder="AED 1,500"
                readOnly={isReadOnly}
              />
            </div>

            {/* Effective Date - only show for resigned/terminated */}
            {(formData.employment?.status === 'resigned' || formData.employment?.status === 'terminated') && (
              <div className="md:col-span-2 lg:col-span-3">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Effective Date <span className="text-red-500">*</span>
                </label>
                <input
                  type="date"
                  value={formData.employment?.effectiveDate || ''}
                  onChange={(e) => handleInputChange('employment', 'effectiveDate', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                  required
                  readOnly={isReadOnly}
                />
                <p className="text-sm text-gray-500 mt-1">
                  Date when the {formData.employment?.status} became effective
                </p>
              </div>
            )}
          </div>
        </div>

        {/* Philippine Recruitment Agency */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-yellow-100 rounded-xl mr-4">
              <Users className="h-6 w-6 text-yellow-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Philippine Recruitment Agency</h2>
              <p className="text-gray-600">Agency that recruited the housemaid from Philippines</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Agency Name <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={formData.recruitmentAgency?.name || ''}
                onChange={(e) => handleInputChange('recruitmentAgency', 'name', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                required
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                License Number
              </label>
              <input
                type="text"
                value={formData.recruitmentAgency?.licenseNumber || ''}
                onChange={(e) => handleInputChange('recruitmentAgency', 'licenseNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Contact Person
              </label>
              <input
                type="text"
                value={formData.recruitmentAgency?.contactPerson || ''}
                onChange={(e) => handleInputChange('recruitmentAgency', 'contactPerson', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Phone Number
              </label>
              <input
                type="tel"
                value={formData.recruitmentAgency?.phoneNumber || ''}
                onChange={(e) => handleInputChange('recruitmentAgency', 'phoneNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Email
              </label>
              <input
                type="email"
                value={formData.recruitmentAgency?.email || ''}
                onChange={(e) => handleInputChange('recruitmentAgency', 'email', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Address
              </label>
              <textarea
                value={formData.recruitmentAgency?.address || ''}
                onChange={(e) => handleInputChange('recruitmentAgency', 'address', e.target.value)}
                rows={3}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Saudi Recruitment Agency */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-green-100 rounded-xl mr-4">
              <Users className="h-6 w-6 text-green-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Saudi Recruitment Agency</h2>
              <p className="text-gray-600">Local Saudi agency handling the placement</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Agency Name
              </label>
              <input
                type="text"
                value={formData.saudiRecruitmentAgency?.name || ''}
                onChange={(e) => handleInputChange('saudiRecruitmentAgency', 'name', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                License Number
              </label>
              <input
                type="text"
                value={formData.saudiRecruitmentAgency?.licenseNumber || ''}
                onChange={(e) => handleInputChange('saudiRecruitmentAgency', 'licenseNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Contact Person
              </label>
              <input
                type="text"
                value={formData.saudiRecruitmentAgency?.contactPerson || ''}
                onChange={(e) => handleInputChange('saudiRecruitmentAgency', 'contactPerson', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Phone Number
              </label>
              <input
                type="tel"
                value={formData.saudiRecruitmentAgency?.phoneNumber || ''}
                onChange={(e) => handleInputChange('saudiRecruitmentAgency', 'phoneNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Email
              </label>
              <input
                type="email"
                value={formData.saudiRecruitmentAgency?.email || ''}
                onChange={(e) => handleInputChange('saudiRecruitmentAgency', 'email', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Address
              </label>
              <textarea
                value={formData.saudiRecruitmentAgency?.address || ''}
                onChange={(e) => handleInputChange('saudiRecruitmentAgency', 'address', e.target.value)}
                rows={3}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                readOnly={isReadOnly}
              />
            </div>
          </div>
        </div>

        {/* Complaint Information */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-orange-100 rounded-xl mr-4">
              <AlertTriangle className="h-6 w-6 text-orange-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Complaint Information</h2>
              <p className="text-gray-600">Any complaints or issues reported</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Complaint Status
              </label>
              <select
                value={formData.complaint?.status || 'pending'}
                onChange={(e) => handleInputChange('complaint', 'status', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                disabled={isReadOnly}
              >
                <option value="pending">Pending</option>
                <option value="complete">Complete</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Date Reported
              </label>
              <input
                type="date"
                value={formData.complaint?.dateReported || ''}
                onChange={(e) => handleInputChange('complaint', 'dateReported', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                readOnly={isReadOnly}
              />
            </div>

            <div className="md:col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Complaint Description
              </label>
              <textarea
                value={formData.complaint?.description || ''}
                onChange={(e) => handleInputChange('complaint', 'description', e.target.value)}
                rows={4}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                placeholder="Describe any complaints or issues..."
                readOnly={isReadOnly}
              />
            </div>

            {formData.complaint?.status === 'complete' && (
              <>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Date Resolved
                  </label>
                  <input
                    type="date"
                    value={formData.complaint?.dateResolved || ''}
                    onChange={(e) => handleInputChange('complaint', 'dateResolved', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                    readOnly={isReadOnly}
                  />
                </div>

                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Resolution Description
                  </label>
                  <textarea
                    value={formData.complaint?.resolutionDescription || ''}
                    onChange={(e) => handleInputChange('complaint', 'resolutionDescription', e.target.value)}
                    rows={4}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                    placeholder="Describe how the complaint was resolved..."
                    readOnly={isReadOnly}
                  />
                </div>
              </>
            )}
          </div>
        </div>

        {/* CV Upload */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-blue-100 rounded-xl mr-4">
              <FileText className="h-6 w-6 text-blue-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">CV/Resume</h2>
              <p className="text-gray-600">Upload curriculum vitae or resume document</p>
            </div>
          </div>

          <CVUpload
            cv={formData.cv!}
            onUpload={handleCVUpload}
            onRemove={handleCVRemove}
            onView={() => setShowCVViewer(true)}
          />
        </div>

        {/* POLO Clearance */}
        <div className="bg-white rounded-xl shadow-sm border p-8">
          <div className="flex items-center mb-6">
            <div className="p-3 bg-emerald-100 rounded-xl mr-4">
              <Shield className="h-6 w-6 text-emerald-600" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">POLO Clearance</h2>
              <p className="text-gray-600">Philippine Overseas Labor Office clearance document</p>
            </div>
          </div>

          <POLOClearanceUpload
            clearance={formData.poloClearance!}
            onUpload={handlePOLOUpload}
            onRemove={handlePOLORemove}
            onView={() => setShowPOLOViewer(true)}
            completionDate={formData.poloClearance?.completionDate}
            onCompletionDateChange={handlePOLOCompletionDateChange}
          />
        </div>

        {/* Form Actions */}
        <div className="flex justify-end space-x-4 pt-8">
          <button
            type="button"
            onClick={onCancel}
            className="px-8 py-3 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors font-medium"
          >
            {isReadOnly ? 'Close' : 'Cancel'}
          </button>
          {!isReadOnly && (
            <button
              type="submit"
              className="px-8 py-3 bg-gradient-to-r from-blue-600 to-purple-600 text-white rounded-lg hover:from-blue-700 hover:to-purple-700 transition-all duration-200 font-medium flex items-center space-x-2 shadow-sm hover:shadow-md"
            >
              <Save className="h-5 w-5" />
              <span>{housemaid ? 'Update Record' : 'Save Record'}</span>
            </button>
          )}
        </div>
      </form>

      {/* Photo Viewer Modal */}
      {showPhotoViewer && formData.profilePhoto?.fileData && (
        <ProfilePhotoViewer
          photo={formData.profilePhoto}
          housemaidName={formData.personalInfo?.name || 'Unknown'}
          onClose={() => setShowPhotoViewer(false)}
        />
      )}

      {/* CV Viewer Modal */}
      {showCVViewer && formData.cv?.fileData && (
        <CVViewer
          cv={formData.cv}
          onClose={() => setShowCVViewer(false)}
        />
      )}

      {/* POLO Clearance Viewer Modal */}
      {showPOLOViewer && formData.poloClearance?.fileData && (
        <POLOClearanceViewer
          clearance={formData.poloClearance}
          housemaidName={formData.personalInfo?.name || 'Unknown'}
          onClose={() => setShowPOLOViewer(false)}
        />
      )}

      {/* Air Ticket Viewer Modal */}
      {showAirTicketViewer && formData.airTicket?.fileData && (
        <AirTicketViewer
          airTicket={formData.airTicket}
          housemaidName={formData.personalInfo?.name || 'Unknown'}
          onClose={() => setShowAirTicketViewer(false)}
        />
      )}
    </div>
  );
};

export default HousemaidForm;