import React, { useState, useEffect } from 'react';
import { Save, X, User, Phone, MapPin, FileText, Building, Plane, Briefcase, Camera, CreditCard, Hash, AlertCircle, Shield, Users, Ticket } from 'lucide-react';
import { Housemaid } from '../types/housemaid';
import { passportCountries } from '../data/countries';
import { generateHousemaidNumberIfEligible, shouldGenerateHousemaidNumber } from '../utils/localStorage';
import CountrySelect from './CountrySelect';
import CVUpload from './CVUpload';
import CVViewer from './CVViewer';
import ProfilePhotoUpload from './ProfilePhotoUpload';
import ProfilePhotoViewer from './ProfilePhotoViewer';
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
  const [formData, setFormData] = useState<Housemaid>({
    id: '',
    housemaidNumber: '',
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
    createdAt: '',
    updatedAt: ''
  });

  const [showCVViewer, setShowCVViewer] = useState(false);
  const [showPhotoViewer, setShowPhotoViewer] = useState(false);
  const [showPOLOViewer, setShowPOLOViewer] = useState(false);
  const [showAirTicketViewer, setShowAirTicketViewer] = useState(false);

  useEffect(() => {
    if (housemaid) {
      // Handle migration from old format to new format
      const migratedData = {
        ...housemaid,
        housemaidNumber: housemaid.housemaidNumber || (housemaid.personalInfo.name.trim() ? generateHousemaidNumberIfEligible(housemaid.personalInfo.name) : ''),
        personalInfo: {
          ...housemaid.personalInfo,
          email: housemaid.personalInfo.email || '',
          citizenship: housemaid.personalInfo.citizenship || '',
          phone: housemaid.personalInfo.phone || 
                 (housemaid.personalInfo as any).countryCode && (housemaid.personalInfo as any).mobileNumber
                   ? `${(housemaid.personalInfo as any).countryCode} ${(housemaid.personalInfo as any).mobileNumber}`
                   : '',
          country: housemaid.personalInfo.country || '',
          city: housemaid.personalInfo.city || ''
        },
        profilePhoto: housemaid.profilePhoto || {
          fileName: undefined,
          fileData: undefined,
          fileType: undefined,
          uploadDate: undefined
        },
        identity: housemaid.identity || {
          passportNumber: (housemaid as any).passport?.number || '',
          passportCountry: (housemaid as any).passport?.country || '',
          residentId: ''
        },
        locationStatus: {
          ...housemaid.locationStatus,
          outsideCountryDate: housemaid.locationStatus.outsideCountryDate || undefined
        },
        flightInfo: {
          ...housemaid.flightInfo,
          flightDate: housemaid.flightInfo?.flightDate || undefined,
          flightNumber: housemaid.flightInfo?.flightNumber || '',
          airlineName: housemaid.flightInfo?.airlineName || '',
          destination: housemaid.flightInfo?.destination || ''
        },
        airTicket: housemaid.airTicket || {
          fileName: undefined,
          fileData: undefined,
          fileType: undefined,
          uploadDate: undefined,
          ticketNumber: '',
          bookingReference: ''
        },
        employment: {
          ...housemaid.employment,
          contractPeriodYears: housemaid.employment?.contractPeriodYears || 2,
          startDate: housemaid.employment?.startDate || '',
          endDate: housemaid.employment?.endDate || '',
          status: housemaid.employment?.status || 'probationary',
          position: housemaid.employment?.position || '',
          salary: housemaid.employment?.salary || '',
          effectiveDate: housemaid.employment?.effectiveDate || undefined
        },
        recruitmentAgency: housemaid.recruitmentAgency || {
          name: '',
          licenseNumber: '',
          contactPerson: '',
          phoneNumber: '',
          email: '',
          address: ''
        },
        saudiRecruitmentAgency: housemaid.saudiRecruitmentAgency || {
          name: '',
          licenseNumber: '',
          contactPerson: '',
          phoneNumber: '',
          email: '',
          address: ''
        },
        complaint: {
          ...housemaid.complaint,
          resolutionDescription: housemaid.complaint.resolutionDescription || ''
        },
        cv: housemaid.cv || {
          fileName: undefined,
          fileData: undefined,
          fileType: undefined,
          uploadDate: undefined
        },
        poloClearance: housemaid.poloClearance || {
          fileName: undefined,
          fileData: undefined,
          fileType: undefined,
          uploadDate: undefined,
          completionDate: undefined
        }
      };
      setFormData(migratedData);
    }
  }, [housemaid]);

  // Auto-calculate end date when start date or contract period changes
  useEffect(() => {
    if (formData.employment.startDate && formData.employment.contractPeriodYears) {
      const startDate = new Date(formData.employment.startDate);
      const endDate = new Date(startDate);
      endDate.setFullYear(startDate.getFullYear() + formData.employment.contractPeriodYears);
      
      setFormData(prev => ({
        ...prev,
        employment: {
          ...prev.employment,
          endDate: endDate.toISOString().split('T')[0]
        }
      }));
    }
  }, [formData.employment.startDate, formData.employment.contractPeriodYears]);

  // Generate housemaid number when name is entered (for new records only)
  useEffect(() => {
    if (!housemaid && shouldGenerateHousemaidNumber(formData.personalInfo.name, formData.housemaidNumber)) {
      const newNumber = generateHousemaidNumberIfEligible(formData.personalInfo.name);
      if (newNumber) {
        setFormData(prev => ({
          ...prev,
          housemaidNumber: newNumber
        }));
      }
    }
  }, [formData.personalInfo.name, housemaid]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    // Validate that name is provided
    if (!formData.personalInfo.name.trim()) {
      alert('Full name is required to create a housemaid record.');
      return;
    }

    // Validate that recruitment agency name is provided
    if (!formData.recruitmentAgency.name.trim()) {
      alert('Philippine Recruitment Agency name is required.');
      return;
    }

    // Generate housemaid number if not already present
    let finalHousemaidNumber = formData.housemaidNumber;
    if (!finalHousemaidNumber && formData.personalInfo.name.trim()) {
      finalHousemaidNumber = generateHousemaidNumberIfEligible(formData.personalInfo.name);
    }

    const now = new Date().toISOString();
    const submissionData = {
      ...formData,
      id: formData.id || Date.now().toString(),
      housemaidNumber: finalHousemaidNumber,
      createdAt: formData.createdAt || now,
      updatedAt: now,
      complaint: {
        ...formData.complaint,
        dateResolved: formData.complaint.status === 'complete' && !formData.complaint.dateResolved
          ? new Date().toISOString().split('T')[0]
          : formData.complaint.dateResolved
      }
    };
    onSave(submissionData);
  };

  const updatePersonalInfo = (field: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      personalInfo: { ...prev.personalInfo, [field]: value }
    }));
  };

  const updateIdentity = (field: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      identity: { ...prev.identity, [field]: value }
    }));
  };

  const updateLocationStatus = (field: string, value: any) => {
    setFormData(prev => ({
      ...prev,
      locationStatus: { ...prev.locationStatus, [field]: value }
    }));
  };

  const updateFlightInfo = (field: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      flightInfo: { ...prev.flightInfo, [field]: value }
    }));
  };

  const updateEmployer = (field: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      employer: { ...prev.employer, [field]: value }
    }));
  };

  const updateEmployment = (field: string, value: any) => {
    setFormData(prev => ({
      ...prev,
      employment: { ...prev.employment, [field]: value }
    }));
  };

  const updateRecruitmentAgency = (field: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      recruitmentAgency: { ...prev.recruitmentAgency, [field]: value }
    }));
  };

  const updateSaudiRecruitmentAgency = (field: string, value: string) => {
    setFormData(prev => ({
      ...prev,
      saudiRecruitmentAgency: { ...prev.saudiRecruitmentAgency, [field]: value }
    }));
  };

  const updateComplaint = (field: string, value: any) => {
    setFormData(prev => ({
      ...prev,
      complaint: { ...prev.complaint, [field]: value }
    }));
  };

  const handleProfilePhotoUpload = (photoData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => {
    setFormData(prev => ({
      ...prev,
      profilePhoto: photoData
    }));
  };

  const handleProfilePhotoRemove = () => {
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

  const handleProfilePhotoView = () => {
    setShowPhotoViewer(true);
  };

  const handleCVUpload = (cvData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => {
    setFormData(prev => ({
      ...prev,
      cv: cvData
    }));
  };

  const handleCVRemove = () => {
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

  const handleCVView = () => {
    setShowCVViewer(true);
  };

  const handlePOLOClearanceUpload = (clearanceData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => {
    setFormData(prev => ({
      ...prev,
      poloClearance: {
        ...prev.poloClearance,
        ...clearanceData
      }
    }));
  };

  const handlePOLOClearanceRemove = () => {
    setFormData(prev => ({
      ...prev,
      poloClearance: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined,
        completionDate: prev.poloClearance.completionDate // Keep completion date
      }
    }));
  };

  const handlePOLOClearanceView = () => {
    setShowPOLOViewer(true);
  };

  const handlePOLOCompletionDateChange = (date: string) => {
    setFormData(prev => ({
      ...prev,
      poloClearance: {
        ...prev.poloClearance,
        completionDate: date
      }
    }));
  };

  const handleAirTicketUpload = (ticketData: {
    fileName: string;
    fileData: string;
    fileType: string;
    uploadDate: string;
  }) => {
    setFormData(prev => ({
      ...prev,
      airTicket: {
        ...prev.airTicket,
        ...ticketData
      }
    }));
  };

  const handleAirTicketRemove = () => {
    setFormData(prev => ({
      ...prev,
      airTicket: {
        fileName: undefined,
        fileData: undefined,
        fileType: undefined,
        uploadDate: undefined,
        ticketNumber: prev.airTicket.ticketNumber, // Keep ticket details
        bookingReference: prev.airTicket.bookingReference
      }
    }));
  };

  const handleAirTicketView = () => {
    setShowAirTicketViewer(true);
  };

  const handleTicketNumberChange = (ticketNumber: string) => {
    setFormData(prev => ({
      ...prev,
      airTicket: {
        ...prev.airTicket,
        ticketNumber
      }
    }));
  };

  const handleBookingReferenceChange = (bookingReference: string) => {
    setFormData(prev => ({
      ...prev,
      airTicket: {
        ...prev.airTicket,
        bookingReference
      }
    }));
  };

  const calculateRemainingDays = () => {
    if (!formData.employment.endDate) return null;
    const today = new Date();
    const endDate = new Date(formData.employment.endDate);
    const diffTime = endDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  };

  const remainingDays = calculateRemainingDays();

  // Check if effective date should be shown
  const shouldShowEffectiveDate = formData.employment.status === 'resigned' || formData.employment.status === 'terminated';

  return (
    <div className="max-w-6xl mx-auto bg-white rounded-xl shadow-lg p-8">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h2 className="text-3xl font-bold text-gray-900">
            {housemaid ? 'Edit Housemaid Record' : 'Add New Housemaid'}
          </h2>
          {formData.housemaidNumber && (
            <div className="flex items-center mt-2">
              <Hash className="h-5 w-5 text-blue-600 mr-2" />
              <span className="text-lg font-semibold text-blue-600 bg-blue-50 px-3 py-1 rounded-full">
                {formData.housemaidNumber}
              </span>
            </div>
          )}
        </div>
        <button
          onClick={onCancel}
          className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
        >
          <X className="h-6 w-6 text-gray-500" />
        </button>
      </div>

      <form onSubmit={handleSubmit} className="space-y-8">
        {/* Housemaid Number Section */}
        <div className={`p-6 rounded-lg border ${formData.housemaidNumber ? 'bg-blue-50 border-blue-200' : 'bg-yellow-50 border-yellow-200'}`}>
          <div className="flex items-center mb-4">
            <Hash className={`h-5 w-5 mr-2 ${formData.housemaidNumber ? 'text-blue-600' : 'text-yellow-600'}`} />
            <h3 className={`text-xl font-semibold ${formData.housemaidNumber ? 'text-blue-900' : 'text-yellow-900'}`}>
              Housemaid Number
            </h3>
          </div>
          <div className={`p-4 rounded-lg border-2 ${formData.housemaidNumber ? 'bg-white border-blue-300' : 'bg-yellow-100 border-yellow-300'}`}>
            {formData.housemaidNumber ? (
              <div className="text-center">
                <p className="text-sm text-blue-700 mb-2">Unique Identification Number</p>
                <p className="text-3xl font-bold text-blue-900 tracking-wider">
                  {formData.housemaidNumber}
                </p>
                <p className="text-xs text-blue-600 mt-2">
                  This number is automatically generated and cannot be changed
                </p>
              </div>
            ) : (
              <div className="text-center">
                <div className="flex items-center justify-center mb-2">
                  <AlertCircle className="h-5 w-5 text-yellow-600 mr-2" />
                  <p className="text-sm font-medium text-yellow-800">Housemaid Number Not Generated</p>
                </div>
                <p className="text-yellow-700 text-sm">
                  A unique housemaid number will be automatically generated once you enter the full name below.
                </p>
                <p className="text-xs text-yellow-600 mt-2">
                  The full name is required to create a unique identification number
                </p>
              </div>
            )}
          </div>
        </div>

        {/* Profile Photo Section */}
        <div className="bg-gradient-to-r from-purple-50 to-pink-50 p-6 rounded-lg border border-purple-200">
          <div className="flex items-center mb-4">
            <Camera className="h-5 w-5 text-purple-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Profile Photo</h3>
          </div>
          <ProfilePhotoUpload
            photo={formData.profilePhoto}
            onUpload={handleProfilePhotoUpload}
            onRemove={handleProfilePhotoRemove}
            onView={handleProfilePhotoView}
          />
        </div>

        {/* Personal Information */}
        <div className="bg-gradient-to-r from-blue-50 to-indigo-50 p-6 rounded-lg border border-blue-200">
          <div className="flex items-center mb-4">
            <User className="h-5 w-5 text-blue-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Personal Information</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Full Name <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                value={formData.personalInfo.name}
                onChange={(e) => updatePersonalInfo('name', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="Enter full name (required for housemaid number)"
                required
              />
              {!formData.personalInfo.name.trim() && (
                <p className="text-xs text-red-600 mt-1">
                  Full name is required to generate a housemaid number
                </p>
              )}
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Email Address</label>
              <input
                type="email"
                value={formData.personalInfo.email}
                onChange={(e) => updatePersonalInfo('email', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="Enter email address"
                required
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Citizenship</label>
              <CountrySelect
                value={formData.personalInfo.citizenship}
                onChange={(value) => updatePersonalInfo('citizenship', value)}
                options={passportCountries.map(country => ({ name: country }))}
                placeholder="Select citizenship country"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Phone Number (with Country Code)</label>
              <input
                type="tel"
                value={formData.personalInfo.phone}
                onChange={(e) => updatePersonalInfo('phone', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="e.g., +971 50 123 4567"
                required
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Country</label>
              <CountrySelect
                value={formData.personalInfo.country}
                onChange={(value) => updatePersonalInfo('country', value)}
                options={passportCountries.map(country => ({ name: country }))}
                placeholder="Select country"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">City</label>
              <input
                type="text"
                value={formData.personalInfo.city}
                onChange={(e) => updatePersonalInfo('city', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="Enter city name"
                required
              />
            </div>
            <div className="md:col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">Address</label>
              <textarea
                value={formData.personalInfo.address}
                onChange={(e) => updatePersonalInfo('address', e.target.value)}
                rows={3}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                placeholder="Enter full address"
                required
              />
            </div>
          </div>

          {/* CV Upload within Personal Information */}
          <div className="border-t border-blue-200 pt-6">
            <div className="flex items-center mb-4">
              <FileText className="h-5 w-5 text-blue-600 mr-2" />
              <h4 className="text-lg font-semibold text-gray-900">CV / Resume</h4>
            </div>
            <CVUpload
              cv={formData.cv}
              onUpload={handleCVUpload}
              onRemove={handleCVRemove}
              onView={handleCVView}
            />
          </div>
        </div>

        {/* Identity Information */}
        <div className="bg-gradient-to-r from-green-50 to-emerald-50 p-6 rounded-lg border border-green-200">
          <div className="flex items-center mb-4">
            <CreditCard className="h-5 w-5 text-green-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Identity Information</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Passport Number</label>
              <input
                type="text"
                value={formData.identity.passportNumber}
                onChange={(e) => updateIdentity('passportNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                placeholder="Enter passport number"
                required
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Passport Country</label>
              <CountrySelect
                value={formData.identity.passportCountry}
                onChange={(value) => updateIdentity('passportCountry', value)}
                options={passportCountries.map(country => ({ name: country }))}
                placeholder="Select passport country"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Resident ID</label>
              <input
                type="text"
                value={formData.identity.residentId || ''}
                onChange={(e) => updateIdentity('residentId', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                placeholder="Enter resident ID number"
              />
            </div>
          </div>
        </div>

        {/* Location Status */}
        <div className="bg-gradient-to-r from-orange-50 to-red-50 p-6 rounded-lg border border-orange-200">
          <div className="flex items-center mb-4">
            <MapPin className="h-5 w-5 text-orange-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Location Status</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Current Location</label>
              <select
                value={formData.locationStatus.isInsideCountry ? 'inside' : 'outside'}
                onChange={(e) => updateLocationStatus('isInsideCountry', e.target.value === 'inside')}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
              >
                <option value="inside">Inside Country</option>
                <option value="outside">Outside Country</option>
              </select>
            </div>
            {!formData.locationStatus.isInsideCountry && (
              <>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Exit Date</label>
                  <input
                    type="date"
                    value={formData.locationStatus.exitDate || ''}
                    onChange={(e) => updateLocationStatus('exitDate', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Date Outside Country</label>
                  <input
                    type="date"
                    value={formData.locationStatus.outsideCountryDate || ''}
                    onChange={(e) => updateLocationStatus('outsideCountryDate', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500 focus:border-orange-500"
                  />
                </div>
              </>
            )}
          </div>
        </div>

        {/* Employer Information */}
        <div className="bg-gradient-to-r from-gray-50 to-slate-50 p-6 rounded-lg border border-gray-200">
          <div className="flex items-center mb-4">
            <Building className="h-5 w-5 text-gray-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Employer Information</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Employer Name</label>
              <input
                type="text"
                value={formData.employer.name}
                onChange={(e) => updateEmployer('name', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-gray-500 focus:border-gray-500"
                placeholder="Enter employer name"
                required
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Employer Mobile</label>
              <input
                type="tel"
                value={formData.employer.mobileNumber}
                onChange={(e) => updateEmployer('mobileNumber', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-gray-500 focus:border-gray-500"
                placeholder="Enter employer mobile number"
                required
              />
            </div>
          </div>
        </div>

        {/* Employment Information */}
        <div className="bg-gradient-to-r from-indigo-50 to-purple-50 p-6 rounded-lg border border-indigo-200">
          <div className="flex items-center mb-4">
            <Briefcase className="h-5 w-5 text-indigo-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Employment Information</h3>
          </div>
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Employment Status</label>
                <select
                  value={formData.employment.status}
                  onChange={(e) => updateEmployment('status', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  required
                >
                  <option value="probationary">Probationary</option>
                  <option value="permanent">Permanent</option>
                  <option value="resigned">Resigned</option>
                  <option value="terminated">Terminated</option>
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Contract Period (Years)</label>
                <select
                  value={formData.employment.contractPeriodYears}
                  onChange={(e) => updateEmployment('contractPeriodYears', parseInt(e.target.value))}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  required
                >
                  <option value={1}>1 Year</option>
                  <option value={2}>2 Years</option>
                  <option value={3}>3 Years</option>
                  <option value={4}>4 Years</option>
                  <option value={5}>5 Years</option>
                </select>
              </div>
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Employment Start Date</label>
                <input
                  type="date"
                  value={formData.employment.startDate}
                  onChange={(e) => updateEmployment('startDate', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  required
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Employment End Date</label>
                <input
                  type="date"
                  value={formData.employment.endDate}
                  onChange={(e) => updateEmployment('endDate', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 bg-gray-100"
                  readOnly
                />
                {remainingDays !== null && (
                  <p className="text-sm text-gray-600 mt-1">
                    {remainingDays > 0 
                      ? `${remainingDays} days remaining` 
                      : remainingDays === 0 
                        ? 'Contract expires today' 
                        : `Contract expired ${Math.abs(remainingDays)} days ago`
                    }
                  </p>
                )}
              </div>
            </div>

            {/* Effective Date for Resigned/Terminated */}
            {shouldShowEffectiveDate && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Effective Date <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="date"
                    value={formData.employment.effectiveDate || ''}
                    onChange={(e) => updateEmployment('effectiveDate', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                    required={shouldShowEffectiveDate}
                  />
                  <p className="text-xs text-gray-600 mt-1">
                    Date when {formData.employment.status} status became effective
                  </p>
                </div>
              </div>
            )}

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Position / Job Title</label>
                <input
                  type="text"
                  value={formData.employment.position || ''}
                  onChange={(e) => updateEmployment('position', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  placeholder="e.g., Housemaid, Domestic Helper"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Monthly Salary</label>
                <input
                  type="text"
                  value={formData.employment.salary || ''}
                  onChange={(e) => updateEmployment('salary', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                  placeholder="e.g., AED 1,500"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Flight Information */}
        <div className="bg-gradient-to-r from-purple-50 to-pink-50 p-6 rounded-lg border border-purple-200">
          <div className="flex items-center mb-4">
            <Plane className="h-5 w-5 text-purple-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Flight Information</h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Flight Date</label>
              <input
                type="date"
                value={formData.flightInfo.flightDate || ''}
                onChange={(e) => updateFlightInfo('flightDate', e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Flight Number</label>
              <input
                type="text"
                value={formData.flightInfo.flightNumber || ''}
                onChange={(e) => updateFlightInfo('flightNumber', e.target.value)}
                placeholder="e.g., EK123"
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Airline Name</label>
              <input
                type="text"
                value={formData.flightInfo.airlineName || ''}
                onChange={(e) => updateFlightInfo('airlineName', e.target.value)}
                placeholder="e.g., Emirates, Philippine Airlines"
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Destination</label>
              <input
                type="text"
                value={formData.flightInfo.destination || ''}
                onChange={(e) => updateFlightInfo('destination', e.target.value)}
                placeholder="e.g., Dubai, Manila"
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-purple-500"
              />
            </div>
          </div>

          {/* Air Ticket Upload within Flight Information */}
          <div className="border-t border-purple-200 pt-6">
            <div className="flex items-center mb-4">
              <Ticket className="h-5 w-5 text-purple-600 mr-2" />
              <h4 className="text-lg font-semibold text-gray-900">Air Ticket</h4>
            </div>
            <AirTicketUpload
              airTicket={formData.airTicket}
              onUpload={handleAirTicketUpload}
              onRemove={handleAirTicketRemove}
              onView={handleAirTicketView}
              ticketNumber={formData.airTicket.ticketNumber}
              bookingReference={formData.airTicket.bookingReference}
              onTicketNumberChange={handleTicketNumberChange}
              onBookingReferenceChange={handleBookingReferenceChange}
            />
          </div>
        </div>

        {/* Licensed Recruitment Agencies */}
        <div className="bg-gradient-to-r from-yellow-50 to-orange-50 p-6 rounded-lg border border-yellow-200">
          <div className="flex items-center mb-6">
            <Users className="h-5 w-5 text-yellow-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Licensed Recruitment Agencies</h3>
          </div>

          {/* Philippine Recruitment Agency */}
          <div className="bg-yellow-100 p-6 rounded-lg border border-yellow-300 mb-6">
            <div className="flex items-center mb-4">
              <Users className="h-5 w-5 text-yellow-700 mr-2" />
              <h4 className="text-lg font-semibold text-gray-900">Philippine Recruitment Agency</h4>
              <span className="ml-2 text-xs bg-red-100 text-red-800 px-2 py-1 rounded-full">Required</span>
            </div>
            <div className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Agency Name <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={formData.recruitmentAgency.name}
                    onChange={(e) => updateRecruitmentAgency('name', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                    placeholder="Enter recruitment agency name"
                    required
                  />
                  {!formData.recruitmentAgency.name.trim() && (
                    <p className="text-xs text-red-600 mt-1">
                      Recruitment agency name is required
                    </p>
                  )}
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">License Number</label>
                  <input
                    type="text"
                    value={formData.recruitmentAgency.licenseNumber || ''}
                    onChange={(e) => updateRecruitmentAgency('licenseNumber', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                    placeholder="Enter agency license number"
                  />
                </div>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Contact Person</label>
                  <input
                    type="text"
                    value={formData.recruitmentAgency.contactPerson || ''}
                    onChange={(e) => updateRecruitmentAgency('contactPerson', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                    placeholder="Enter contact person name"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Phone Number</label>
                  <input
                    type="tel"
                    value={formData.recruitmentAgency.phoneNumber || ''}
                    onChange={(e) => updateRecruitmentAgency('phoneNumber', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                    placeholder="Enter agency phone number"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Email Address</label>
                  <input
                    type="email"
                    value={formData.recruitmentAgency.email || ''}
                    onChange={(e) => updateRecruitmentAgency('email', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                    placeholder="Enter agency email address"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Address</label>
                  <textarea
                    value={formData.recruitmentAgency.address || ''}
                    onChange={(e) => updateRecruitmentAgency('address', e.target.value)}
                    rows={3}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-yellow-500 focus:border-yellow-500"
                    placeholder="Enter agency address"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Saudi Recruitment Agency */}
          <div className="bg-green-100 p-6 rounded-lg border border-green-300">
            <div className="flex items-center mb-4">
              <Users className="h-5 w-5 text-green-700 mr-2" />
              <h4 className="text-lg font-semibold text-gray-900">Saudi Recruitment Agency</h4>
              <span className="ml-2 text-xs bg-blue-100 text-blue-800 px-2 py-1 rounded-full">Optional</span>
            </div>
            <div className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Agency Name</label>
                  <input
                    type="text"
                    value={formData.saudiRecruitmentAgency.name}
                    onChange={(e) => updateSaudiRecruitmentAgency('name', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                    placeholder="Enter Saudi recruitment agency name"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">License Number</label>
                  <input
                    type="text"
                    value={formData.saudiRecruitmentAgency.licenseNumber || ''}
                    onChange={(e) => updateSaudiRecruitmentAgency('licenseNumber', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                    placeholder="Enter agency license number"
                  />
                </div>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Contact Person</label>
                  <input
                    type="text"
                    value={formData.saudiRecruitmentAgency.contactPerson || ''}
                    onChange={(e) => updateSaudiRecruitmentAgency('contactPerson', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                    placeholder="Enter contact person name"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Phone Number</label>
                  <input
                    type="tel"
                    value={formData.saudiRecruitmentAgency.phoneNumber || ''}
                    onChange={(e) => updateSaudiRecruitmentAgency('phoneNumber', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                    placeholder="Enter agency phone number"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Email Address</label>
                  <input
                    type="email"
                    value={formData.saudiRecruitmentAgency.email || ''}
                    onChange={(e) => updateSaudiRecruitmentAgency('email', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                    placeholder="Enter agency email address"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Address</label>
                  <textarea
                    value={formData.saudiRecruitmentAgency.address || ''}
                    onChange={(e) => updateSaudiRecruitmentAgency('address', e.target.value)}
                    rows={3}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                    placeholder="Enter agency address"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Complaint Information */}
        <div className="bg-gradient-to-r from-red-50 to-pink-50 p-6 rounded-lg border border-red-200">
          <div className="flex items-center mb-4">
            <Phone className="h-5 w-5 text-red-600 mr-2" />
            <h3 className="text-xl font-semibold text-gray-900">Complaint Information</h3>
          </div>
          <div className="space-y-6 mb-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Complaint Description</label>
              <textarea
                value={formData.complaint.description}
                onChange={(e) => updateComplaint('description', e.target.value)}
                rows={4}
                className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-red-500"
                placeholder="Describe the complaint or issue..."
              />
            </div>
            
            {formData.complaint.status === 'complete' && (
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Resolution Description</label>
                <textarea
                  value={formData.complaint.resolutionDescription || ''}
                  onChange={(e) => updateComplaint('resolutionDescription', e.target.value)}
                  rows={4}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500 bg-green-50"
                  placeholder="Describe how the complaint was resolved..."
                />
              </div>
            )}
            
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Status</label>
                <select
                  value={formData.complaint.status}
                  onChange={(e) => updateComplaint('status', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-red-500"
                >
                  <option value="pending">Pending</option>
                  <option value="complete">Complete</option>
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Date Reported</label>
                <input
                  type="date"
                  value={formData.complaint.dateReported}
                  onChange={(e) => updateComplaint('dateReported', e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-red-500 focus:border-red-500"
                  required
                />
              </div>
              {formData.complaint.status === 'complete' && (
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Date Resolved</label>
                  <input
                    type="date"
                    value={formData.complaint.dateResolved || ''}
                    onChange={(e) => updateComplaint('dateResolved', e.target.value)}
                    className="w-full px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
                  />
                </div>
              )}
            </div>
          </div>

          {/* POLO Clearance Upload within Complaint Information */}
          <div className="border-t border-red-200 pt-6">
            <div className="flex items-center mb-4">
              <Shield className="h-5 w-5 text-red-600 mr-2" />
              <h4 className="text-lg font-semibold text-gray-900">POLO Clearance</h4>
            </div>
            <POLOClearanceUpload
              clearance={formData.poloClearance}
              onUpload={handlePOLOClearanceUpload}
              onRemove={handlePOLOClearanceRemove}
              onView={handlePOLOClearanceView}
              completionDate={formData.poloClearance.completionDate}
              onCompletionDateChange={handlePOLOCompletionDateChange}
            />
          </div>
        </div>

        {/* Form Actions */}
        <div className="flex justify-end space-x-4 pt-6 border-t">
          <button
            type="button"
            onClick={onCancel}
            className="px-6 py-3 text-gray-700 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors font-medium"
          >
            Cancel
          </button>
          <button
            type="submit"
            className="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium flex items-center space-x-2"
          >
            <Save className="h-5 w-5" />
            <span>{housemaid ? 'Update Record' : 'Save Record'}</span>
          </button>
        </div>
      </form>

      {/* CV Viewer Modal */}
      {showCVViewer && (
        <CVViewer
          cv={formData.cv}
          onClose={() => setShowCVViewer(false)}
        />
      )}

      {/* Profile Photo Viewer Modal */}
      {showPhotoViewer && (
        <ProfilePhotoViewer
          photo={formData.profilePhoto}
          housemaidName={formData.personalInfo.name}
          onClose={() => setShowPhotoViewer(false)}
        />
      )}

      {/* POLO Clearance Viewer Modal */}
      {showPOLOViewer && (
        <POLOClearanceViewer
          clearance={formData.poloClearance}
          housemaidName={formData.personalInfo.name}
          onClose={() => setShowPOLOViewer(false)}
        />
      )}

      {/* Air Ticket Viewer Modal */}
      {showAirTicketViewer && (
        <AirTicketViewer
          airTicket={formData.airTicket}
          housemaidName={formData.personalInfo.name}
          onClose={() => setShowAirTicketViewer(false)}
        />
      )}
    </div>
  );
};

export default HousemaidForm;