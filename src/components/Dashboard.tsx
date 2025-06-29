import React, { useState, useMemo } from 'react';
import { 
  Users, 
  AlertTriangle, 
  CheckCircle, 
  MapPin, 
  Plane, 
  FileText, 
  Clock, 
  Briefcase, 
  Camera, 
  Shield,
  TrendingUp,
  Calendar,
  Phone,
  Mail,
  Eye,
  X,
  Search,
  Hash,
  Building2,
  Ticket,
  Award,
  UserX,
  UserMinus
} from 'lucide-react';
import { Housemaid } from '../types/housemaid';
import { calculateEmploymentStatus, formatDaysWorked, getEmploymentStatusColor, getEmploymentStatusText, getDisplayDaysWorked } from '../utils/employmentCalculations';
import StatusBadge from './StatusBadge';

interface DashboardProps {
  housemaids: Housemaid[];
  onViewHousemaid: (housemaid: Housemaid) => void;
}

interface DashboardCard {
  title: string;
  value: number;
  icon: React.ReactNode;
  color: string;
  bgColor: string;
  data: Housemaid[];
  description: string;
}

interface RecruitmentAgencyStats {
  name: string;
  count: number;
  housemaids: Housemaid[];
  percentage: number;
}

const Dashboard: React.FC<DashboardProps> = ({ housemaids, onViewHousemaid }) => {
  const [selectedCard, setSelectedCard] = useState<DashboardCard | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedAgency, setSelectedAgency] = useState<RecruitmentAgencyStats | null>(null);

  const dashboardData = useMemo(() => {
    const today = new Date();
    const thirtyDaysFromNow = new Date();
    thirtyDaysFromNow.setDate(today.getDate() + 30);

    const pendingComplaints = housemaids.filter(h => h.complaint.status === 'pending');
    const resolvedComplaints = housemaids.filter(h => h.complaint.status === 'complete' && h.complaint.resolutionDescription);
    const outsideCountry = housemaids.filter(h => !h.locationStatus.isInsideCountry);
    const upcomingFlights = housemaids.filter(h => {
      if (!h.flightInfo?.flightDate) return false;
      const flightDate = new Date(h.flightInfo.flightDate);
      return flightDate >= today;
    });
    const withCV = housemaids.filter(h => h.cv?.fileName && h.cv?.fileData);
    const expiringContracts = housemaids.filter(h => {
      if (!h.employment?.endDate) return false;
      const endDate = new Date(h.employment.endDate);
      return endDate >= today && endDate <= thirtyDaysFromNow;
    });
    const probationary = housemaids.filter(h => h.employment?.status === 'probationary');
    const permanent = housemaids.filter(h => h.employment?.status === 'permanent');
    const resigned = housemaids.filter(h => h.employment?.status === 'resigned');
    const terminated = housemaids.filter(h => h.employment?.status === 'terminated');
    const withPhotos = housemaids.filter(h => h.profilePhoto?.fileName && h.profilePhoto?.fileData);
    const withPOLO = housemaids.filter(h => h.poloClearance?.fileName && h.poloClearance?.fileData);
    const withAirTickets = housemaids.filter(h => h.airTicket?.fileName && h.airTicket?.fileData);

    // Calculate eligible for permanent (90+ days but still probationary)
    const eligibleForPermanent = housemaids.filter(h => {
      if (h.employment?.status !== 'probationary' || !h.employment?.startDate) return false;
      const calculation = calculateEmploymentStatus(h.employment.startDate, h.employment.status, h.employment.effectiveDate);
      return calculation.isEligibleForPermanent;
    });

    return [
      {
        title: 'Total Records',
        value: housemaids.length,
        icon: <Users className="h-8 w-8" />,
        color: 'text-blue-600',
        bgColor: 'bg-blue-50 border-blue-200',
        data: housemaids,
        description: 'All housemaid records in the system'
      },
      {
        title: 'Pending Complaints',
        value: pendingComplaints.length,
        icon: <AlertTriangle className="h-8 w-8" />,
        color: 'text-orange-600',
        bgColor: 'bg-orange-50 border-orange-200',
        data: pendingComplaints,
        description: 'Complaints that need attention'
      },
      {
        title: 'Resolved Issues',
        value: resolvedComplaints.length,
        icon: <CheckCircle className="h-8 w-8" />,
        color: 'text-green-600',
        bgColor: 'bg-green-50 border-green-200',
        data: resolvedComplaints,
        description: 'Successfully resolved complaints'
      },
      {
        title: 'Outside Country',
        value: outsideCountry.length,
        icon: <MapPin className="h-8 w-8" />,
        color: 'text-red-600',
        bgColor: 'bg-red-50 border-red-200',
        data: outsideCountry,
        description: 'Currently outside the country'
      },
      {
        title: 'Upcoming Flights',
        value: upcomingFlights.length,
        icon: <Plane className="h-8 w-8" />,
        color: 'text-purple-600',
        bgColor: 'bg-purple-50 border-purple-200',
        data: upcomingFlights,
        description: 'Scheduled flights from today onwards'
      },
      {
        title: 'CVs Uploaded',
        value: withCV.length,
        icon: <FileText className="h-8 w-8" />,
        color: 'text-indigo-600',
        bgColor: 'bg-indigo-50 border-indigo-200',
        data: withCV,
        description: 'Records with CV documents'
      },
      {
        title: 'Contracts Expiring',
        value: expiringContracts.length,
        icon: <Clock className="h-8 w-8" />,
        color: 'text-yellow-600',
        bgColor: 'bg-yellow-50 border-yellow-200',
        data: expiringContracts,
        description: 'Contracts expiring within 30 days'
      },
      {
        title: 'Probationary',
        value: probationary.length,
        icon: <Briefcase className="h-8 w-8" />,
        color: 'text-amber-600',
        bgColor: 'bg-amber-50 border-amber-200',
        data: probationary,
        description: 'Employees in probationary period'
      },
      {
        title: 'Permanent Staff',
        value: permanent.length,
        icon: <Award className="h-8 w-8" />,
        color: 'text-emerald-600',
        bgColor: 'bg-emerald-50 border-emerald-200',
        data: permanent,
        description: 'Permanent employees'
      },
      {
        title: 'Resigned',
        value: resigned.length,
        icon: <UserMinus className="h-8 w-8" />,
        color: 'text-blue-600',
        bgColor: 'bg-blue-50 border-blue-200',
        data: resigned,
        description: 'Employees who have resigned'
      },
      {
        title: 'Terminated',
        value: terminated.length,
        icon: <UserX className="h-8 w-8" />,
        color: 'text-red-600',
        bgColor: 'bg-red-50 border-red-200',
        data: terminated,
        description: 'Terminated employees'
      },
      {
        title: 'Eligible for Permanent',
        value: eligibleForPermanent.length,
        icon: <TrendingUp className="h-8 w-8" />,
        color: 'text-cyan-600',
        bgColor: 'bg-cyan-50 border-cyan-200',
        data: eligibleForPermanent,
        description: 'Probationary staff eligible for permanent status (90+ days)'
      },
      {
        title: 'Profile Photos',
        value: withPhotos.length,
        icon: <Camera className="h-8 w-8" />,
        color: 'text-pink-600',
        bgColor: 'bg-pink-50 border-pink-200',
        data: withPhotos,
        description: 'Records with profile photos'
      },
      {
        title: 'POLO Clearance',
        value: withPOLO.length,
        icon: <Shield className="h-8 w-8" />,
        color: 'text-green-600',
        bgColor: 'bg-green-50 border-green-200',
        data: withPOLO,
        description: 'Records with POLO clearance documents'
      },
      {
        title: 'Air Tickets',
        value: withAirTickets.length,
        icon: <Ticket className="h-8 w-8" />,
        color: 'text-violet-600',
        bgColor: 'bg-violet-50 border-violet-200',
        data: withAirTickets,
        description: 'Records with air ticket documents'
      }
    ];
  }, [housemaids]);

  const recruitmentAgencyStats = useMemo(() => {
    const agencyMap = new Map<string, Housemaid[]>();
    
    // Group housemaids by recruitment agency
    housemaids.forEach(housemaid => {
      const agencyName = housemaid.recruitmentAgency?.name?.trim();
      if (agencyName) {
        if (!agencyMap.has(agencyName)) {
          agencyMap.set(agencyName, []);
        }
        agencyMap.get(agencyName)!.push(housemaid);
      }
    });

    // Convert to stats array and sort by count
    const stats: RecruitmentAgencyStats[] = Array.from(agencyMap.entries())
      .map(([name, housemaids]) => ({
        name,
        count: housemaids.length,
        housemaids,
        percentage: Math.round((housemaids.length / housemaids.length) * 100)
      }))
      .sort((a, b) => b.count - a.count);

    // Calculate correct percentages
    const totalWithAgency = stats.reduce((sum, stat) => sum + stat.count, 0);
    stats.forEach(stat => {
      stat.percentage = totalWithAgency > 0 ? Math.round((stat.count / totalWithAgency) * 100) : 0;
    });

    return stats;
  }, [housemaids]);

  const filteredData = useMemo(() => {
    if (!selectedCard || !searchTerm) return selectedCard?.data || [];
    
    return selectedCard.data.filter(housemaid => {
      const searchLower = searchTerm.toLowerCase();
      return (
        housemaid.personalInfo.name.toLowerCase().includes(searchLower) ||
        housemaid.personalInfo.phone.includes(searchTerm) ||
        (housemaid.personalInfo.email && housemaid.personalInfo.email.toLowerCase().includes(searchLower)) ||
        (housemaid.housemaidNumber && housemaid.housemaidNumber.toLowerCase().includes(searchLower)) ||
        housemaid.employer.name.toLowerCase().includes(searchLower)
      );
    });
  }, [selectedCard, searchTerm]);

  const formatDate = (dateString: string) => {
    return new Date(dateString).toLocaleDateString();
  };

  const calculateRemainingDays = (endDate: string) => {
    const today = new Date();
    const end = new Date(endDate);
    const diffTime = end.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  };

  const getRecentActivity = () => {
    const sortedByUpdated = [...housemaids].sort((a, b) => 
      new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime()
    );
    return sortedByUpdated.slice(0, 5);
  };

  const recentActivity = getRecentActivity();

  return (
    <div className="space-y-8">
      {/* Header */}
      <div className="bg-gradient-to-r from-blue-600 to-purple-600 rounded-2xl p-8 text-white">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-4xl font-bold mb-2">Dashboard Overview</h1>
            <p className="text-blue-100 text-lg">
              Comprehensive view of your housemaid management system
            </p>
          </div>
          <div className="text-right">
            <div className="text-3xl font-bold">{housemaids.length}</div>
            <div className="text-blue-100">Total Records</div>
          </div>
        </div>
      </div>

      {/* Dashboard Cards Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-6">
        {dashboardData.map((card, index) => (
          <div
            key={index}
            onClick={() => setSelectedCard(card)}
            className={`${card.bgColor} border-2 rounded-xl p-6 cursor-pointer hover:shadow-lg transition-all duration-300 hover:scale-105 group`}
          >
            <div className="flex items-center justify-between mb-4">
              <div className={`${card.color} group-hover:scale-110 transition-transform duration-300`}>
                {card.icon}
              </div>
              <Eye className="h-5 w-5 text-gray-400 group-hover:text-gray-600 transition-colors" />
            </div>
            <div className="space-y-2">
              <h3 className="font-semibold text-gray-900 text-sm">{card.title}</h3>
              <div className="text-3xl font-bold text-gray-900">{card.value}</div>
              <p className="text-xs text-gray-600">{card.description}</p>
            </div>
          </div>
        ))}
      </div>

      {/* Philippine Recruitment Agency Statistics */}
      {recruitmentAgencyStats.length > 0 && (
        <div className="bg-white rounded-xl shadow-sm border p-6">
          <div className="flex items-center justify-between mb-6">
            <h2 className="text-2xl font-bold text-gray-900 flex items-center">
              <Building2 className="h-6 w-6 mr-2 text-yellow-600" />
              Philippine Recruitment Agencies
            </h2>
            <div className="text-sm text-gray-500">
              {recruitmentAgencyStats.length} {recruitmentAgencyStats.length === 1 ? 'agency' : 'agencies'} • {recruitmentAgencyStats.reduce((sum, stat) => sum + stat.count, 0)} total recruits
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {recruitmentAgencyStats.map((agency, index) => (
              <div
                key={agency.name}
                onClick={() => setSelectedAgency(agency)}
                className="bg-yellow-50 border border-yellow-200 rounded-lg p-4 hover:shadow-md cursor-pointer transition-all duration-200 hover:bg-yellow-100 group"
              >
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center space-x-2">
                    <div className="w-8 h-8 bg-yellow-600 rounded-lg flex items-center justify-center text-white font-bold text-sm">
                      {index + 1}
                    </div>
                    <div className="text-yellow-600">
                      <Building2 className="h-5 w-5" />
                    </div>
                  </div>
                  <Eye className="h-4 w-4 text-gray-400 group-hover:text-yellow-600 transition-colors" />
                </div>
                
                <div className="space-y-2">
                  <h4 className="font-semibold text-gray-900 text-sm line-clamp-2 group-hover:text-yellow-800 transition-colors">
                    {agency.name}
                  </h4>
                  <div className="flex items-center justify-between">
                    <span className="text-2xl font-bold text-yellow-700">{agency.count}</span>
                    <span className="text-sm text-yellow-600 font-medium">{agency.percentage}%</span>
                  </div>
                  <p className="text-xs text-yellow-600">
                    {agency.count === 1 ? 'housemaid recruited' : 'housemaids recruited'}
                  </p>
                </div>

                {/* Progress bar */}
                <div className="mt-3">
                  <div className="w-full bg-yellow-200 rounded-full h-2">
                    <div 
                      className="bg-yellow-600 h-2 rounded-full transition-all duration-300"
                      style={{ width: `${agency.percentage}%` }}
                    ></div>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* Summary Stats */}
          <div className="mt-6 pt-6 border-t border-gray-200">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="text-center">
                <div className="text-2xl font-bold text-gray-900">{recruitmentAgencyStats.length}</div>
                <div className="text-sm text-gray-600">Active Agencies</div>
              </div>
              <div className="text-center">
                <div className="text-2xl font-bold text-gray-900">
                  {recruitmentAgencyStats.reduce((sum, stat) => sum + stat.count, 0)}
                </div>
                <div className="text-sm text-gray-600">Total Recruited</div>
              </div>
              <div className="text-center">
                <div className="text-2xl font-bold text-gray-900">
                  {recruitmentAgencyStats.length > 0 ? Math.round(recruitmentAgencyStats.reduce((sum, stat) => sum + stat.count, 0) / recruitmentAgencyStats.length) : 0}
                </div>
                <div className="text-sm text-gray-600">Average per Agency</div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Recent Activity */}
      <div className="bg-white rounded-xl shadow-sm border p-6">
        <div className="flex items-center justify-between mb-6">
          <h2 className="text-2xl font-bold text-gray-900 flex items-center">
            <TrendingUp className="h-6 w-6 mr-2 text-blue-600" />
            Recent Activity
          </h2>
          <span className="text-sm text-gray-500">Last 5 updates</span>
        </div>
        <div className="space-y-4">
          {recentActivity.map((housemaid) => {
            const employmentCalc = housemaid.employment?.startDate 
              ? calculateEmploymentStatus(housemaid.employment.startDate, housemaid.employment.status, housemaid.employment.effectiveDate)
              : null;

            const displayDays = employmentCalc ? getDisplayDaysWorked(employmentCalc) : 0;

            return (
              <div
                key={housemaid.id}
                onClick={() => onViewHousemaid(housemaid)}
                className="flex items-center justify-between p-4 bg-gray-50 rounded-lg hover:bg-gray-100 cursor-pointer transition-colors group"
              >
                <div className="flex items-center space-x-4">
                  {housemaid.profilePhoto?.fileData ? (
                    <img
                      src={housemaid.profilePhoto.fileData}
                      alt={housemaid.personalInfo.name}
                      className="w-12 h-12 rounded-full object-cover border-2 border-gray-200"
                    />
                  ) : (
                    <div className="w-12 h-12 rounded-full bg-gray-200 flex items-center justify-center">
                      <Users className="h-6 w-6 text-gray-400" />
                    </div>
                  )}
                  <div>
                    <h4 className="font-semibold text-gray-900 group-hover:text-blue-600 transition-colors">
                      {housemaid.personalInfo.name}
                    </h4>
                    <div className="flex items-center space-x-2 text-sm text-gray-600">
                      {housemaid.housemaidNumber && (
                        <span className="flex items-center">
                          <Hash className="h-3 w-3 mr-1" />
                          {housemaid.housemaidNumber}
                        </span>
                      )}
                      <span>•</span>
                      <span>Updated {formatDate(housemaid.updatedAt)}</span>
                      {housemaid.recruitmentAgency?.name && (
                        <>
                          <span>•</span>
                          <span className="text-yellow-600 font-medium">{housemaid.recruitmentAgency.name}</span>
                        </>
                      )}
                      {employmentCalc && (
                        <>
                          <span>•</span>
                          <span className="text-blue-600 font-medium">
                            {formatDaysWorked(displayDays)} worked
                            {employmentCalc.isResignedOrTerminated && ' (total)'}
                          </span>
                        </>
                      )}
                    </div>
                  </div>
                </div>
                <div className="flex items-center space-x-2">
                  {housemaid.employment?.status && employmentCalc && (
                    <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium border ${getEmploymentStatusColor(housemaid.employment.status, displayDays)}`}>
                      {getEmploymentStatusText(housemaid.employment.status, displayDays)}
                    </span>
                  )}
                  <StatusBadge
                    status={housemaid.complaint.status}
                    text={housemaid.complaint.status === 'pending' ? 'Pending' : 'Complete'}
                  />
                  <Eye className="h-4 w-4 text-gray-400 group-hover:text-blue-600 transition-colors" />
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Quick Stats */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
        <div className="bg-gradient-to-br from-green-50 to-emerald-50 border border-green-200 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <CheckCircle className="h-8 w-8 text-green-600" />
            <span className="text-2xl font-bold text-green-600">
              {Math.round((dashboardData[2].value / housemaids.length) * 100) || 0}%
            </span>
          </div>
          <h3 className="font-semibold text-gray-900 mb-2">Resolution Rate</h3>
          <p className="text-sm text-gray-600">Percentage of resolved complaints</p>
        </div>

        <div className="bg-gradient-to-br from-blue-50 to-indigo-50 border border-blue-200 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <FileText className="h-8 w-8 text-blue-600" />
            <span className="text-2xl font-bold text-blue-600">
              {Math.round((dashboardData[5].value / housemaids.length) * 100) || 0}%
            </span>
          </div>
          <h3 className="font-semibold text-gray-900 mb-2">Documentation Rate</h3>
          <p className="text-sm text-gray-600">Records with complete CV documentation</p>
        </div>

        <div className="bg-gradient-to-br from-purple-50 to-pink-50 border border-purple-200 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <Camera className="h-8 w-8 text-purple-600" />
            <span className="text-2xl font-bold text-purple-600">
              {Math.round((dashboardData[12].value / housemaids.length) * 100) || 0}%
            </span>
          </div>
          <h3 className="font-semibold text-gray-900 mb-2">Profile Completion</h3>
          <p className="text-sm text-gray-600">Records with profile photos</p>
        </div>

        <div className="bg-gradient-to-br from-violet-50 to-purple-50 border border-violet-200 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <Ticket className="h-8 w-8 text-violet-600" />
            <span className="text-2xl font-bold text-violet-600">
              {Math.round((dashboardData[14].value / housemaids.length) * 100) || 0}%
            </span>
          </div>
          <h3 className="font-semibold text-gray-900 mb-2">Air Tickets</h3>
          <p className="text-sm text-gray-600">Records with air ticket documents</p>
        </div>
      </div>

      {/* Detail Modal */}
      {selectedCard && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-6xl w-full max-h-[90vh] overflow-hidden shadow-2xl">
            {/* Modal Header */}
            <div className={`${selectedCard.bgColor} border-b p-6`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center space-x-4">
                  <div className={`${selectedCard.color} p-3 bg-white rounded-xl shadow-sm`}>
                    {selectedCard.icon}
                  </div>
                  <div>
                    <h3 className="text-2xl font-bold text-gray-900">{selectedCard.title}</h3>
                    <p className="text-gray-600">{selectedCard.description}</p>
                    <div className="flex items-center space-x-4 mt-2">
                      <span className="text-3xl font-bold text-gray-900">{selectedCard.value}</span>
                      <span className="text-sm text-gray-500">
                        {selectedCard.value === 1 ? 'record' : 'records'}
                      </span>
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => setSelectedCard(null)}
                  className="p-2 hover:bg-white hover:bg-opacity-20 rounded-lg transition-colors"
                >
                  <X className="h-6 w-6 text-gray-600" />
                </button>
              </div>
            </div>

            {/* Search Bar */}
            {selectedCard.data.length > 0 && (
              <div className="p-6 border-b bg-gray-50">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-5 w-5 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search records..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full pl-10 pr-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white"
                  />
                </div>
              </div>
            )}

            {/* Modal Content */}
            <div className="p-6 overflow-y-auto max-h-[calc(90vh-200px)]">
              {filteredData.length === 0 ? (
                <div className="text-center py-12">
                  <div className={`${selectedCard.color} mx-auto mb-4`}>
                    {selectedCard.icon}
                  </div>
                  <h4 className="text-lg font-medium text-gray-900 mb-2">
                    {searchTerm ? 'No matching records found' : 'No records in this category'}
                  </h4>
                  <p className="text-gray-600">
                    {searchTerm ? 'Try adjusting your search terms.' : selectedCard.description}
                  </p>
                </div>
              ) : (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  {filteredData.map((housemaid) => {
                    const employmentCalc = housemaid.employment?.startDate 
                      ? calculateEmploymentStatus(housemaid.employment.startDate, housemaid.employment.status, housemaid.employment.effectiveDate)
                      : null;

                    const displayDays = employmentCalc ? getDisplayDaysWorked(employmentCalc) : 0;

                    return (
                      <div
                        key={housemaid.id}
                        className="bg-white border border-gray-200 rounded-xl p-6 hover:shadow-lg transition-all duration-300 hover:border-blue-300 group cursor-pointer"
                        onClick={() => {
                          setSelectedCard(null);
                          onViewHousemaid(housemaid);
                        }}
                      >
                        <div className="flex items-start justify-between mb-4">
                          <div className="flex items-center space-x-4">
                            {housemaid.profilePhoto?.fileData ? (
                              <img
                                src={housemaid.profilePhoto.fileData}
                                alt={housemaid.personalInfo.name}
                                className="w-16 h-16 rounded-full object-cover border-2 border-gray-200 group-hover:border-blue-400 transition-colors"
                              />
                            ) : (
                              <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center border-2 border-gray-200 group-hover:border-blue-400 transition-colors">
                                <Users className="h-8 w-8 text-gray-400" />
                              </div>
                            )}
                            <div>
                              <h4 className="font-semibold text-gray-900 group-hover:text-blue-600 transition-colors">
                                {housemaid.personalInfo.name}
                              </h4>
                              {housemaid.housemaidNumber && (
                                <div className="flex items-center mt-1">
                                  <Hash className="h-3 w-3 text-blue-600 mr-1" />
                                  <span className="text-sm font-medium text-blue-600">
                                    {housemaid.housemaidNumber}
                                  </span>
                                </div>
                              )}
                              <div className="flex items-center text-sm text-gray-600 mt-1">
                                <Phone className="h-3 w-3 mr-1" />
                                {housemaid.personalInfo.phone}
                              </div>
                              {housemaid.personalInfo.email && (
                                <div className="flex items-center text-sm text-gray-600 mt-1">
                                  <Mail className="h-3 w-3 mr-1" />
                                  {housemaid.personalInfo.email}
                                </div>
                              )}
                            </div>
                          </div>
                          <Eye className="h-5 w-5 text-gray-400 group-hover:text-blue-600 transition-colors" />
                        </div>

                        <div className="space-y-3">
                          <div className="flex items-center justify-between">
                            <span className="text-sm text-gray-600">Employer:</span>
                            <span className="text-sm font-medium">{housemaid.employer.name}</span>
                          </div>

                          {housemaid.recruitmentAgency?.name && (
                            <div className="flex items-center justify-between">
                              <span className="text-sm text-gray-600">Agency:</span>
                              <span className="text-sm font-medium text-yellow-700">{housemaid.recruitmentAgency.name}</span>
                            </div>
                          )}

                          <div className="flex items-center justify-between">
                            <span className="text-sm text-gray-600">Status:</span>
                            <StatusBadge
                              status={housemaid.complaint.status}
                              text={housemaid.complaint.status === 'pending' ? 'Pending' : 'Complete'}
                            />
                          </div>

                          <div className="flex items-center justify-between">
                            <span className="text-sm text-gray-600">Location:</span>
                            <StatusBadge
                              status={housemaid.locationStatus.isInsideCountry ? 'inside' : 'outside'}
                              text={housemaid.locationStatus.isInsideCountry ? 'Inside Country' : 'Outside Country'}
                            />
                          </div>

                          {housemaid.employment?.status && employmentCalc && (
                            <div className="flex items-center justify-between">
                              <span className="text-sm text-gray-600">Employment:</span>
                              <div className="text-right">
                                <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium border ${getEmploymentStatusColor(housemaid.employment.status, displayDays)}`}>
                                  {getEmploymentStatusText(housemaid.employment.status, displayDays)}
                                </span>
                                <div className="text-xs text-gray-500 mt-1">
                                  {formatDaysWorked(displayDays)} worked
                                  {employmentCalc.isResignedOrTerminated && ' (total)'}
                                </div>
                              </div>
                            </div>
                          )}

                          {housemaid.employment?.effectiveDate && (housemaid.employment.status === 'resigned' || housemaid.employment.status === 'terminated') && (
                            <div className="flex items-center justify-between">
                              <span className="text-sm text-gray-600">Effective Date:</span>
                              <div className="text-right">
                                <div className="text-sm font-medium">{formatDate(housemaid.employment.effectiveDate)}</div>
                                <div className="text-xs text-gray-500">
                                  {housemaid.employment.status === 'resigned' ? 'Resignation' : 'Termination'} date
                                </div>
                              </div>
                            </div>
                          )}

                          {housemaid.employment?.endDate && (
                            <div className="flex items-center justify-between">
                              <span className="text-sm text-gray-600">Contract Ends:</span>
                              <div className="text-right">
                                <div className="text-sm font-medium">{formatDate(housemaid.employment.endDate)}</div>
                                {(() => {
                                  const remainingDays = calculateRemainingDays(housemaid.employment.endDate);
                                  return (
                                    <div className={`text-xs ${
                                      remainingDays <= 30 && remainingDays > 0 ? 'text-yellow-600' :
                                      remainingDays <= 0 ? 'text-red-600' : 'text-gray-500'
                                    }`}>
                                      {remainingDays > 0 
                                        ? `${remainingDays} days left` 
                                        : remainingDays === 0 
                                          ? 'Expires today' 
                                          : `Expired ${Math.abs(remainingDays)} days ago`
                                      }
                                    </div>
                                  );
                                })()}
                              </div>
                            </div>
                          )}

                          {housemaid.flightInfo?.flightDate && (
                            <div className="flex items-center justify-between">
                              <span className="text-sm text-gray-600">Flight Date:</span>
                              <div className="flex items-center text-sm font-medium">
                                <Calendar className="h-3 w-3 mr-1 text-purple-600" />
                                {formatDate(housemaid.flightInfo.flightDate)}
                              </div>
                            </div>
                          )}

                          {housemaid.airTicket?.ticketNumber && (
                            <div className="flex items-center justify-between">
                              <span className="text-sm text-gray-600">Ticket:</span>
                              <div className="flex items-center text-sm font-medium">
                                <Ticket className="h-3 w-3 mr-1 text-violet-600" />
                                {housemaid.airTicket.ticketNumber}
                              </div>
                            </div>
                          )}

                          {housemaid.complaint.description && (
                            <div className="pt-3 border-t">
                              <p className="text-sm text-gray-600 mb-1">Latest Complaint:</p>
                              <p className="text-sm text-gray-800 line-clamp-2">{housemaid.complaint.description}</p>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Agency Detail Modal */}
      {selectedAgency && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl max-w-4xl w-full max-h-[90vh] overflow-hidden shadow-2xl">
            {/* Modal Header */}
            <div className="bg-yellow-50 border-b border-yellow-200 p-6">
              <div className="flex items-center justify-between">
                <div className="flex items-center space-x-4">
                  <div className="p-3 bg-yellow-600 rounded-xl shadow-sm">
                    <Building2 className="h-6 w-6 text-white" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-bold text-gray-900">{selectedAgency.name}</h3>
                    <p className="text-gray-600">Philippine Recruitment Agency</p>
                    <div className="flex items-center space-x-4 mt-2">
                      <span className="text-3xl font-bold text-yellow-700">{selectedAgency.count}</span>
                      <span className="text-sm text-gray-500">
                        {selectedAgency.count === 1 ? 'housemaid recruited' : 'housemaids recruited'}
                      </span>
                      <span className="text-sm font-medium text-yellow-600">({selectedAgency.percentage}%)</span>
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => setSelectedAgency(null)}
                  className="p-2 hover:bg-yellow-100 rounded-lg transition-colors"
                >
                  <X className="h-6 w-6 text-gray-600" />
                </button>
              </div>
            </div>

            {/* Modal Content */}
            <div className="p-6 overflow-y-auto max-h-[calc(90vh-200px)]">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {selectedAgency.housemaids.map((housemaid) => {
                  const employmentCalc = housemaid.employment?.startDate 
                    ? calculateEmploymentStatus(housemaid.employment.startDate, housemaid.employment.status, housemaid.employment.effectiveDate)
                    : null;

                  const displayDays = employmentCalc ? getDisplayDaysWorked(employmentCalc) : 0;

                  return (
                    <div
                      key={housemaid.id}
                      className="bg-white border border-gray-200 rounded-xl p-6 hover:shadow-lg transition-all duration-300 hover:border-yellow-300 group cursor-pointer"
                      onClick={() => {
                        setSelectedAgency(null);
                        onViewHousemaid(housemaid);
                      }}
                    >
                      <div className="flex items-start justify-between mb-4">
                        <div className="flex items-center space-x-4">
                          {housemaid.profilePhoto?.fileData ? (
                            <img
                              src={housemaid.profilePhoto.fileData}
                              alt={housemaid.personalInfo.name}
                              className="w-16 h-16 rounded-full object-cover border-2 border-gray-200 group-hover:border-yellow-400 transition-colors"
                            />
                          ) : (
                            <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center border-2 border-gray-200 group-hover:border-yellow-400 transition-colors">
                              <Users className="h-8 w-8 text-gray-400" />
                            </div>
                          )}
                          <div>
                            <h4 className="font-semibold text-gray-900 group-hover:text-yellow-700 transition-colors">
                              {housemaid.personalInfo.name}
                            </h4>
                            {housemaid.housemaidNumber && (
                              <div className="flex items-center mt-1">
                                <Hash className="h-3 w-3 text-blue-600 mr-1" />
                                <span className="text-sm font-medium text-blue-600">
                                  {housemaid.housemaidNumber}
                                </span>
                              </div>
                            )}
                            <div className="flex items-center text-sm text-gray-600 mt-1">
                              <Phone className="h-3 w-3 mr-1" />
                              {housemaid.personalInfo.phone}
                            </div>
                          </div>
                        </div>
                        <Eye className="h-5 w-5 text-gray-400 group-hover:text-yellow-600 transition-colors" />
                      </div>

                      <div className="space-y-3">
                        <div className="flex items-center justify-between">
                          <span className="text-sm text-gray-600">Employer:</span>
                          <span className="text-sm font-medium">{housemaid.employer.name}</span>
                        </div>

                        <div className="flex items-center justify-between">
                          <span className="text-sm text-gray-600">Status:</span>
                          <StatusBadge
                            status={housemaid.complaint.status}
                            text={housemaid.complaint.status === 'pending' ? 'Pending' : 'Complete'}
                          />
                        </div>

                        <div className="flex items-center justify-between">
                          <span className="text-sm text-gray-600">Location:</span>
                          <StatusBadge
                            status={housemaid.locationStatus.isInsideCountry ? 'inside' : 'outside'}
                            text={housemaid.locationStatus.isInsideCountry ? 'Inside Country' : 'Outside Country'}
                          />
                        </div>

                        {housemaid.employment?.status && employmentCalc && (
                          <div className="flex items-center justify-between">
                            <span className="text-sm text-gray-600">Employment:</span>
                            <div className="text-right">
                              <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium border ${getEmploymentStatusColor(housemaid.employment.status, displayDays)}`}>
                                {getEmploymentStatusText(housemaid.employment.status, displayDays)}
                              </span>
                              <div className="text-xs text-gray-500 mt-1">
                                {formatDaysWorked(displayDays)} worked
                                {employmentCalc.isResignedOrTerminated && ' (total)'}
                              </div>
                            </div>
                          </div>
                        )}

                        {housemaid.employment?.startDate && (
                          <div className="flex items-center justify-between">
                            <span className="text-sm text-gray-600">Start Date:</span>
                            <span className="text-sm font-medium">{formatDate(housemaid.employment.startDate)}</span>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Dashboard;