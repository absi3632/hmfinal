import React, { useState, useMemo } from 'react';
import { 
  Plus, 
  Search, 
  Filter, 
  Download, 
  Upload, 
  Eye, 
  Edit, 
  Trash2, 
  Users, 
  Phone, 
  Mail, 
  Calendar, 
  Building, 
  Hash, 
  FileText, 
  Camera, 
  Shield, 
  Plane, 
  Ticket,
  ChevronDown,
  Grid3X3,
  List,
  SortAsc,
  SortDesc,
  Clock,
  AlertTriangle,
  CheckCircle
} from 'lucide-react';
import { Housemaid } from '../types/housemaid';
import { calculateEmploymentStatus, formatDaysWorked, getEmploymentStatusColor, getEmploymentStatusText, getDisplayDaysWorked } from '../utils/employmentCalculations';
import { hasPermission, getCurrentUser } from '../utils/auth';
import StatusBadge from './StatusBadge';
import ExcelImport from './ExcelImport';
import ReportGenerator from './ReportGenerator';

interface HousemaidListProps {
  housemaids: Housemaid[];
  onAdd: () => void;
  onEdit: (housemaid: Housemaid) => void;
  onDelete: (id: string) => void;
  onBulkImport: (housemaids: Housemaid[]) => void;
}

type ViewMode = 'grid' | 'list';
type SortField = 'name' | 'housemaidNumber' | 'createdAt' | 'updatedAt' | 'employer' | 'agency' | 'status';
type SortOrder = 'asc' | 'desc';

const HousemaidList: React.FC<HousemaidListProps> = ({ 
  housemaids, 
  onAdd, 
  onEdit, 
  onDelete, 
  onBulkImport 
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'pending' | 'complete'>('all');
  const [locationFilter, setLocationFilter] = useState<'all' | 'inside' | 'outside'>('all');
  const [employmentFilter, setEmploymentFilter] = useState<'all' | 'probationary' | 'permanent' | 'resigned' | 'terminated' | 'eligible'>('all');
  const [agencyFilter, setAgencyFilter] = useState<string>('all');
  const [viewMode, setViewMode] = useState<ViewMode>('grid');
  const [sortField, setSortField] = useState<SortField>('updatedAt');
  const [sortOrder, setSortOrder] = useState<SortOrder>('desc');
  const [showFilters, setShowFilters] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showReports, setShowReports] = useState(false);
  const [selectedHousemaids, setSelectedHousemaids] = useState<Set<string>>(new Set());

  const currentUser = getCurrentUser();

  // Get unique agencies for filter
  const agencies = useMemo(() => {
    const agencySet = new Set<string>();
    housemaids.forEach(h => {
      if (h.recruitmentAgency?.name?.trim()) {
        agencySet.add(h.recruitmentAgency.name.trim());
      }
    });
    return Array.from(agencySet).sort();
  }, [housemaids]);

  // Filter and sort housemaids
  const filteredAndSortedHousemaids = useMemo(() => {
    let filtered = housemaids.filter(housemaid => {
      // Search filter
      const searchLower = searchTerm.toLowerCase();
      const matchesSearch = !searchTerm || (
        housemaid.personalInfo.name.toLowerCase().includes(searchLower) ||
        housemaid.personalInfo.phone.includes(searchTerm) ||
        (housemaid.personalInfo.email && housemaid.personalInfo.email.toLowerCase().includes(searchLower)) ||
        (housemaid.housemaidNumber && housemaid.housemaidNumber.toLowerCase().includes(searchLower)) ||
        housemaid.employer.name.toLowerCase().includes(searchLower) ||
        (housemaid.recruitmentAgency?.name && housemaid.recruitmentAgency.name.toLowerCase().includes(searchLower))
      );

      // Status filter
      const matchesStatus = statusFilter === 'all' || housemaid.complaint.status === statusFilter;

      // Location filter
      const matchesLocation = locationFilter === 'all' || 
        (locationFilter === 'inside' && housemaid.locationStatus.isInsideCountry) ||
        (locationFilter === 'outside' && !housemaid.locationStatus.isInsideCountry);

      // Employment filter
      let matchesEmployment = true;
      if (employmentFilter !== 'all') {
        if (employmentFilter === 'eligible') {
          const calc = housemaid.employment?.startDate 
            ? calculateEmploymentStatus(housemaid.employment.startDate, housemaid.employment.status, housemaid.employment.effectiveDate)
            : null;
          matchesEmployment = !!(calc?.isEligibleForPermanent) && housemaid.employment?.status === 'probationary';
        } else {
          matchesEmployment = housemaid.employment?.status === employmentFilter;
        }
      }

      // Agency filter
      const matchesAgency = agencyFilter === 'all' || housemaid.recruitmentAgency?.name === agencyFilter;

      return matchesSearch && matchesStatus && matchesLocation && matchesEmployment && matchesAgency;
    });

    // Sort
    filtered.sort((a, b) => {
      let aValue: any, bValue: any;

      switch (sortField) {
        case 'name':
          aValue = a.personalInfo.name.toLowerCase();
          bValue = b.personalInfo.name.toLowerCase();
          break;
        case 'housemaidNumber':
          aValue = a.housemaidNumber || '';
          bValue = b.housemaidNumber || '';
          break;
        case 'employer':
          aValue = a.employer.name.toLowerCase();
          bValue = b.employer.name.toLowerCase();
          break;
        case 'agency':
          aValue = (a.recruitmentAgency?.name || '').toLowerCase();
          bValue = (b.recruitmentAgency?.name || '').toLowerCase();
          break;
        case 'status':
          aValue = a.complaint.status;
          bValue = b.complaint.status;
          break;
        case 'createdAt':
          aValue = new Date(a.createdAt).getTime();
          bValue = new Date(b.createdAt).getTime();
          break;
        case 'updatedAt':
        default:
          aValue = new Date(a.updatedAt).getTime();
          bValue = new Date(b.updatedAt).getTime();
          break;
      }

      if (aValue < bValue) return sortOrder === 'asc' ? -1 : 1;
      if (aValue > bValue) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });

    return filtered;
  }, [housemaids, searchTerm, statusFilter, locationFilter, employmentFilter, agencyFilter, sortField, sortOrder]);

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortOrder('asc');
    }
  };

  const handleSelectHousemaid = (id: string) => {
    const newSelected = new Set(selectedHousemaids);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedHousemaids(newSelected);
  };

  const handleSelectAll = () => {
    if (selectedHousemaids.size === filteredAndSortedHousemaids.length) {
      setSelectedHousemaids(new Set());
    } else {
      setSelectedHousemaids(new Set(filteredAndSortedHousemaids.map(h => h.id)));
    }
  };

  const handleBulkDelete = () => {
    if (window.confirm(`Are you sure you want to delete ${selectedHousemaids.size} selected records?`)) {
      selectedHousemaids.forEach(id => onDelete(id));
      setSelectedHousemaids(new Set());
    }
  };

  const formatDate = (dateString: string) => {
    return new Date(dateString).toLocaleDateString();
  };

  const getStatusIcon = (status: string) => {
    return status === 'complete' ? 
      <CheckCircle className="h-4 w-4 text-green-600" /> : 
      <AlertTriangle className="h-4 w-4 text-orange-600" />;
  };

  const HousemaidCard: React.FC<{ housemaid: Housemaid }> = ({ housemaid }) => {
    const employmentCalc = housemaid.employment?.startDate 
      ? calculateEmploymentStatus(housemaid.employment.startDate, housemaid.employment.status, housemaid.employment.effectiveDate)
      : null;

    const displayDays = employmentCalc ? getDisplayDaysWorked(employmentCalc) : 0;
    const isSelected = selectedHousemaids.has(housemaid.id);

    return (
      <div className={`bg-white rounded-xl shadow-sm border-2 transition-all duration-300 hover:shadow-lg hover:border-blue-300 group ${isSelected ? 'border-blue-500 bg-blue-50' : 'border-gray-200'}`}>
        {/* Card Header */}
        <div className="p-6 border-b border-gray-100">
          <div className="flex items-start justify-between">
            <div className="flex items-center space-x-4">
              {/* Selection Checkbox */}
              {hasPermission(currentUser?.role || 'viewer', 'manager') && (
                <input
                  type="checkbox"
                  checked={isSelected}
                  onChange={() => handleSelectHousemaid(housemaid.id)}
                  className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500"
                />
              )}

              {/* Profile Photo */}
              <div className="relative">
                {housemaid.profilePhoto?.fileData ? (
                  <img
                    src={housemaid.profilePhoto.fileData}
                    alt={housemaid.personalInfo.name}
                    className="w-16 h-16 rounded-full object-cover border-3 border-white shadow-md"
                  />
                ) : (
                  <div className="w-16 h-16 rounded-full bg-gradient-to-br from-blue-500 to-purple-500 flex items-center justify-center text-white font-bold text-lg shadow-md">
                    {housemaid.personalInfo.name.split(' ').map(n => n[0]).join('').slice(0, 2)}
                  </div>
                )}
                
                {/* Status Indicator */}
                <div className="absolute -bottom-1 -right-1">
                  {getStatusIcon(housemaid.complaint.status)}
                </div>
              </div>

              {/* Basic Info */}
              <div className="flex-1">
                <div className="flex items-center space-x-2 mb-1">
                  <h3 className="text-lg font-bold text-gray-900 group-hover:text-blue-600 transition-colors">
                    {housemaid.personalInfo.name}
                  </h3>
                  {housemaid.housemaidNumber && (
                    <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
                      <Hash className="h-3 w-3 mr-1" />
                      {housemaid.housemaidNumber}
                    </span>
                  )}
                </div>
                
                <div className="space-y-1">
                  <div className="flex items-center text-sm text-gray-600">
                    <Phone className="h-3 w-3 mr-2" />
                    {housemaid.personalInfo.phone}
                  </div>
                  {housemaid.personalInfo.email && (
                    <div className="flex items-center text-sm text-gray-600">
                      <Mail className="h-3 w-3 mr-2" />
                      {housemaid.personalInfo.email}
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Action Menu */}
            <div className="flex items-center space-x-2">
              <button
                onClick={() => onEdit(housemaid)}
                className="p-2 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                title="View Details"
              >
                <Eye className="h-4 w-4" />
              </button>
              
              {hasPermission(currentUser?.role || 'viewer', 'manager') && (
                <>
                  <button
                    onClick={() => onEdit(housemaid)}
                    className="p-2 text-gray-400 hover:text-green-600 hover:bg-green-50 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                    title="Edit Record"
                  >
                    <Edit className="h-4 w-4" />
                  </button>
                  <button
                    onClick={() => onDelete(housemaid.id)}
                    className="p-2 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors opacity-0 group-hover:opacity-100"
                    title="Delete Record"
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </>
              )}
            </div>
          </div>
        </div>

        {/* Card Body */}
        <div className="p-6 space-y-4">
          {/* Status Badges */}
          <div className="flex flex-wrap gap-2">
            <StatusBadge
              status={housemaid.complaint.status}
              text={housemaid.complaint.status === 'pending' ? 'Pending' : 'Complete'}
            />
            <StatusBadge
              status={housemaid.locationStatus.isInsideCountry ? 'inside' : 'outside'}
              text={housemaid.locationStatus.isInsideCountry ? 'Inside Country' : 'Outside Country'}
            />
            {housemaid.employment?.status && employmentCalc && (
              <span className={`inline-flex items-center px-3 py-1 rounded-full text-sm font-medium border ${getEmploymentStatusColor(housemaid.employment.status, displayDays)}`}>
                {getEmploymentStatusText(housemaid.employment.status, displayDays)}
              </span>
            )}
          </div>

          {/* Key Information */}
          <div className="grid grid-cols-1 gap-3">
            <div className="flex items-center justify-between">
              <span className="text-sm text-gray-600 flex items-center">
                <Building className="h-4 w-4 mr-2" />
                Employer
              </span>
              <span className="text-sm font-medium text-gray-900">{housemaid.employer.name}</span>
            </div>

            {housemaid.recruitmentAgency?.name && (
              <div className="flex items-center justify-between">
                <span className="text-sm text-gray-600 flex items-center">
                  <Users className="h-4 w-4 mr-2" />
                  Agency
                </span>
                <span className="text-sm font-medium text-yellow-700">{housemaid.recruitmentAgency.name}</span>
              </div>
            )}

            {employmentCalc && (
              <div className="flex items-center justify-between">
                <span className="text-sm text-gray-600 flex items-center">
                  <Clock className="h-4 w-4 mr-2" />
                  Days Worked
                </span>
                <span className="text-sm font-medium text-blue-600">
                  {formatDaysWorked(displayDays)}
                  {employmentCalc.isResignedOrTerminated && ' (total)'}
                </span>
              </div>
            )}

            {/* Show effective date for resigned/terminated */}
            {housemaid.employment?.effectiveDate && (housemaid.employment.status === 'resigned' || housemaid.employment.status === 'terminated') && (
              <div className="flex items-center justify-between">
                <span className="text-sm text-gray-600 flex items-center">
                  <Calendar className="h-4 w-4 mr-2" />
                  Effective Date
                </span>
                <span className="text-sm font-medium text-gray-900">{formatDate(housemaid.employment.effectiveDate)}</span>
              </div>
            )}

            <div className="flex items-center justify-between">
              <span className="text-sm text-gray-600 flex items-center">
                <Calendar className="h-4 w-4 mr-2" />
                Last Updated
              </span>
              <span className="text-sm font-medium text-gray-900">{formatDate(housemaid.updatedAt)}</span>
            </div>
          </div>

          {/* Document Status */}
          <div className="flex items-center justify-between pt-3 border-t border-gray-100">
            <span className="text-sm text-gray-600">Documents</span>
            <div className="flex items-center space-x-2">
              {housemaid.cv?.fileName && (
                <div className="p-1 bg-blue-100 rounded" title="CV Uploaded">
                  <FileText className="h-3 w-3 text-blue-600" />
                </div>
              )}
              {housemaid.profilePhoto?.fileName && (
                <div className="p-1 bg-green-100 rounded" title="Photo Uploaded">
                  
                  <Camera className="h-3 w-3 text-green-600" />
                </div>
              )}
              {housemaid.poloClearance?.fileName && (
                <div className="p-1 bg-emerald-100 rounded" title="POLO Clearance">
                  <Shield className="h-3 w-3 text-emerald-600" />
                </div>
              )}
              {housemaid.airTicket?.fileName && (
                <div className="p-1 bg-purple-100 rounded" title="Air Ticket">
                  <Ticket className="h-3 w-3 text-purple-600" />
                </div>
              )}
              {housemaid.flightInfo?.flightDate && (
                <div className="p-1 bg-indigo-100 rounded" title="Flight Info">
                  <Plane className="h-3 w-3 text-indigo-600" />
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const HousemaidRow: React.FC<{ housemaid: Housemaid }> = ({ housemaid }) => {
    const employmentCalc = housemaid.employment?.startDate 
      ? calculateEmploymentStatus(housemaid.employment.startDate, housemaid.employment.status, housemaid.employment.effectiveDate)
      : null;

    const displayDays = employmentCalc ? getDisplayDaysWorked(employmentCalc) : 0;
    const isSelected = selectedHousemaids.has(housemaid.id);

    return (
      <tr className={`hover:bg-gray-50 transition-colors ${isSelected ? 'bg-blue-50' : ''}`}>
        {/* Selection */}
        {hasPermission(currentUser?.role || 'viewer', 'manager') && (
          <td className="px-6 py-4">
            <input
              type="checkbox"
              checked={isSelected}
              onChange={() => handleSelectHousemaid(housemaid.id)}
              className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500"
            />
          </td>
        )}

        {/* Photo & Name */}
        <td className="px-6 py-4">
          <div className="flex items-center space-x-3">
            {housemaid.profilePhoto?.fileData ? (
              <img
                src={housemaid.profilePhoto.fileData}
                alt={housemaid.personalInfo.name}
                className="w-10 h-10 rounded-full object-cover"
              />
            ) : (
              <div className="w-10 h-10 rounded-full bg-gradient-to-br from-blue-500 to-purple-500 flex items-center justify-center text-white font-bold text-sm">
                {housemaid.personalInfo.name.split(' ').map(n => n[0]).join('').slice(0, 2)}
              </div>
            )}
            <div>
              <div className="font-medium text-gray-900">{housemaid.personalInfo.name}</div>
              {housemaid.housemaidNumber && (
                <div className="text-sm text-blue-600 font-medium">#{housemaid.housemaidNumber}</div>
              )}
            </div>
          </div>
        </td>

        {/* Contact */}
        <td className="px-6 py-4">
          <div className="text-sm text-gray-900">{housemaid.personalInfo.phone}</div>
          {housemaid.personalInfo.email && (
            <div className="text-sm text-gray-500">{housemaid.personalInfo.email}</div>
          )}
        </td>

        {/* Employer */}
        <td className="px-6 py-4">
          <div className="text-sm font-medium text-gray-900">{housemaid.employer.name}</div>
        </td>

        {/* Agency */}
        <td className="px-6 py-4">
          <div className="text-sm font-medium text-yellow-700">
            {housemaid.recruitmentAgency?.name || '-'}
          </div>
        </td>

        {/* Employment */}
        <td className="px-6 py-4">
          {housemaid.employment?.status && employmentCalc && (
            <div>
              <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium border ${getEmploymentStatusColor(housemaid.employment.status, displayDays)}`}>
                {getEmploymentStatusText(housemaid.employment.status, displayDays)}
              </span>
              <div className="text-xs text-gray-500 mt-1">
                {formatDaysWorked(displayDays)}
                {employmentCalc.isResignedOrTerminated && ' (total)'}
              </div>
              {/* Show effective date for resigned/terminated */}
              {housemaid.employment.effectiveDate && (housemaid.employment.status === 'resigned' || housemaid.employment.status === 'terminated') && (
                <div className="text-xs text-gray-500 mt-1">
                  Effective: {formatDate(housemaid.employment.effectiveDate)}
                </div>
              )}
            </div>
          )}
        </td>

        {/* Status */}
        <td className="px-6 py-4">
          <div className="flex items-center space-x-2">
            <StatusBadge
              status={housemaid.complaint.status}
              text={housemaid.complaint.status === 'pending' ? 'Pending' : 'Complete'}
            />
            <StatusBadge
              status={housemaid.locationStatus.isInsideCountry ? 'inside' : 'outside'}
              text={housemaid.locationStatus.isInsideCountry ? 'Inside' : 'Outside'}
            />
          </div>
        </td>

        {/* Documents */}
        <td className="px-6 py-4">
          <div className="flex items-center space-x-1">
            {housemaid.cv?.fileName && (
              <div className="p-1 bg-blue-100 rounded" title="CV">
                <FileText className="h-3 w-3 text-blue-600" />
              </div>
            )}
            {housemaid.profilePhoto?.fileName && (
              <div className="p-1 bg-green-100 rounded" title="Photo">
                <Camera className="h-3 w-3 text-green-600" />
              </div>
            )}
            {housemaid.poloClearance?.fileName && (
              <div className="p-1 bg-emerald-100 rounded" title="POLO">
                <Shield className="h-3 w-3 text-emerald-600" />
              </div>
            )}
            {housemaid.airTicket?.fileName && (
              <div className="p-1 bg-purple-100 rounded" title="Ticket">
                <Ticket className="h-3 w-3 text-purple-600" />
              </div>
            )}
          </div>
        </td>

        {/* Updated */}
        <td className="px-6 py-4 text-sm text-gray-500">
          {formatDate(housemaid.updatedAt)}
        </td>

        {/* Actions */}
        <td className="px-6 py-4">
          <div className="flex items-center space-x-2">
            <button
              onClick={() => onEdit(housemaid)}
              className="p-1 text-gray-400 hover:text-blue-600 transition-colors"
              title="View Details"
            >
              <Eye className="h-4 w-4" />
            </button>
            
            {hasPermission(currentUser?.role || 'viewer', 'manager') && (
              <>
                <button
                  onClick={() => onEdit(housemaid)}
                  className="p-1 text-gray-400 hover:text-green-600 transition-colors"
                  title="Edit"
                >
                  <Edit className="h-4 w-4" />
                </button>
                <button
                  onClick={() => onDelete(housemaid.id)}
                  className="p-1 text-gray-400 hover:text-red-600 transition-colors"
                  title="Delete"
                >
                  <Trash2 className="h-4 w-4" />
                </button>
              </>
            )}
          </div>
        </td>
      </tr>
    );
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-gradient-to-r from-blue-600 to-purple-600 rounded-2xl p-8 text-white">
        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center space-y-4 lg:space-y-0">
          <div>
            <h1 className="text-4xl font-bold mb-2">Housemaid Database</h1>
            <p className="text-blue-100 text-lg">
              Manage and track all housemaid records with professional tools
            </p>
            <div className="flex items-center space-x-6 mt-4">
              <div className="text-center">
                <div className="text-2xl font-bold">{housemaids.length}</div>
                <div className="text-blue-100 text-sm">Total Records</div>
              </div>
              <div className="text-center">
                <div className="text-2xl font-bold">{filteredAndSortedHousemaids.length}</div>
                <div className="text-blue-100 text-sm">Filtered Results</div>
              </div>
              {selectedHousemaids.size > 0 && (
                <div className="text-center">
                  <div className="text-2xl font-bold">{selectedHousemaids.size}</div>
                  <div className="text-blue-100 text-sm">Selected</div>
                </div>
              )}
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap gap-3">
            {hasPermission(currentUser?.role || 'viewer', 'manager') && (
              <>
                <button
                  onClick={() => setShowImport(true)}
                  className="bg-white bg-opacity-20 text-white px-6 py-3 rounded-lg hover:bg-opacity-30 transition-all duration-200 font-medium flex items-center space-x-2 backdrop-blur-sm"
                >
                  <Upload className="h-5 w-5" />
                  <span>Import Excel</span>
                </button>
                
                <button
                  onClick={onAdd}
                  className="bg-white text-blue-600 px-6 py-3 rounded-lg hover:bg-gray-50 transition-all duration-200 font-medium flex items-center space-x-2 shadow-sm"
                >
                  <Plus className="h-5 w-5" />
                  <span>Add Record</span>
                </button>
              </>
            )}
            
            <button
              onClick={() => setShowReports(true)}
              className="bg-white bg-opacity-20 text-white px-6 py-3 rounded-lg hover:bg-opacity-30 transition-all duration-200 font-medium flex items-center space-x-2 backdrop-blur-sm"
            >
              <Download className="h-5 w-5" />
              <span>Export</span>
            </button>
          </div>
        </div>
      </div>

      {/* Search and Filters */}
      <div className="bg-white rounded-xl shadow-sm border p-6">
        <div className="flex flex-col lg:flex-row gap-4">
          {/* Search */}
          <div className="flex-1">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-5 w-5 text-gray-400" />
              <input
                type="text"
                placeholder="Search by name, phone, email, ID, employer, or agency..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-gray-50 focus:bg-white transition-colors"
              />
            </div>
          </div>

          {/* View Toggle */}
          <div className="flex items-center bg-gray-100 rounded-lg p-1">
            <button
              onClick={() => setViewMode('grid')}
              className={`p-2 rounded-md transition-colors ${viewMode === 'grid' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-600 hover:text-gray-900'}`}
              title="Grid View"
            >
              <Grid3X3 className="h-5 w-5" />
            </button>
            <button
              onClick={() => setViewMode('list')}
              className={`p-2 rounded-md transition-colors ${viewMode === 'list' ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-600 hover:text-gray-900'}`}
              title="List View"
            >
              <List className="h-5 w-5" />
            </button>
          </div>

          {/* Filter Toggle */}
          <button
            onClick={() => setShowFilters(!showFilters)}
            className={`px-4 py-3 rounded-lg border transition-colors flex items-center space-x-2 ${showFilters ? 'bg-blue-50 border-blue-200 text-blue-700' : 'border-gray-300 text-gray-700 hover:bg-gray-50'}`}
          >
            <Filter className="h-5 w-5" />
            <span>Filters</span>
            <ChevronDown className={`h-4 w-4 transition-transform ${showFilters ? 'rotate-180' : ''}`} />
          </button>
        </div>

        {/* Advanced Filters */}
        {showFilters && (
          <div className="mt-6 pt-6 border-t border-gray-200">
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-6 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Status</label>
                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value as any)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="all">All Status</option>
                  <option value="pending">Pending</option>
                  <option value="complete">Complete</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Location</label>
                <select
                  value={locationFilter}
                  onChange={(e) => setLocationFilter(e.target.value as any)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="all">All Locations</option>
                  <option value="inside">Inside Country</option>
                  <option value="outside">Outside Country</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Employment</label>
                <select
                  value={employmentFilter}
                  onChange={(e) => setEmploymentFilter(e.target.value as any)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="all">All Employment</option>
                  <option value="probationary">Probationary</option>
                  <option value="permanent">Permanent</option>
                  <option value="resigned">Resigned</option>
                  <option value="terminated">Terminated</option>
                  <option value="eligible">Eligible for Permanent</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Agency</label>
                <select
                  value={agencyFilter}
                  onChange={(e) => setAgencyFilter(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="all">All Agencies</option>
                  {agencies.map(agency => (
                    <option key={agency} value={agency}>{agency}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Sort By</label>
                <div className="flex space-x-2">
                  <select
                    value={sortField}
                    onChange={(e) => setSortField(e.target.value as SortField)}
                    className="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    <option value="updatedAt">Last Updated</option>
                    <option value="createdAt">Date Created</option>
                    <option value="name">Name</option>
                    <option value="housemaidNumber">ID Number</option>
                    <option value="employer">Employer</option>
                    <option value="agency">Agency</option>
                    <option value="status">Status</option>
                  </select>
                  <button
                    onClick={() => setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')}
                    className="px-3 py-2 border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
                    title={`Sort ${sortOrder === 'asc' ? 'Descending' : 'Ascending'}`}
                  >
                    {sortOrder === 'asc' ? <SortAsc className="h-4 w-4" /> : <SortDesc className="h-4 w-4" />}
                  </button>
                </div>
              </div>
            </div>

            {/* Clear Filters */}
            <div className="mt-4 flex justify-end">
              <button
                onClick={() => {
                  setSearchTerm('');
                  setStatusFilter('all');
                  setLocationFilter('all');
                  setEmploymentFilter('all');
                  setAgencyFilter('all');
                  setSortField('updatedAt');
                  setSortOrder('desc');
                }}
                className="px-4 py-2 text-gray-600 hover:text-gray-900 transition-colors"
              >
                Clear All Filters
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Bulk Actions */}
      {selectedHousemaids.size > 0 && hasPermission(currentUser?.role || 'viewer', 'manager') && (
        <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-4">
              <span className="text-blue-700 font-medium">
                {selectedHousemaids.size} record{selectedHousemaids.size !== 1 ? 's' : ''} selected
              </span>
              <button
                onClick={() => setSelectedHousemaids(new Set())}
                className="text-blue-600 hover:text-blue-800 transition-colors"
              >
                Clear Selection
              </button>
            </div>
            <div className="flex items-center space-x-2">
              <button
                onClick={handleBulkDelete}
                className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors flex items-center space-x-2"
              >
                <Trash2 className="h-4 w-4" />
                <span>Delete Selected</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Results */}
      {filteredAndSortedHousemaids.length === 0 ? (
        <div className="bg-white rounded-xl shadow-sm border p-12 text-center">
          <Users className="h-16 w-16 text-gray-300 mx-auto mb-4" />
          <h3 className="text-lg font-medium text-gray-900 mb-2">No records found</h3>
          <p className="text-gray-600 mb-6">
            {searchTerm || statusFilter !== 'all' || locationFilter !== 'all' || employmentFilter !== 'all' || agencyFilter !== 'all'
              ? 'Try adjusting your search criteria or filters.'
              : 'Get started by adding your first housemaid record.'
            }
          </p>
          {hasPermission(currentUser?.role || 'viewer', 'manager') && (
            <button
              onClick={onAdd}
              className="bg-blue-600 text-white px-6 py-3 rounded-lg hover:bg-blue-700 transition-colors font-medium flex items-center space-x-2 mx-auto"
            >
              <Plus className="h-5 w-5" />
              <span>Add First Record</span>
            </button>
          )}
        </div>
      ) : (
        <>
          {viewMode === 'grid' ? (
            /* Grid View */
            <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-4 gap-6">
              {filteredAndSortedHousemaids.map((housemaid) => (
                <HousemaidCard key={housemaid.id} housemaid={housemaid} />
              ))}
            </div>
          ) : (
            /* List View */
            <div className="bg-white rounded-xl shadow-sm border overflow-hidden">
              <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-gray-200">
                  <thead className="bg-gray-50">
                    <tr>
                      {hasPermission(currentUser?.role || 'viewer', 'manager') && (
                        <th className="px-6 py-3 text-left">
                          <input
                            type="checkbox"
                            checked={selectedHousemaids.size === filteredAndSortedHousemaids.length && filteredAndSortedHousemaids.length > 0}
                            onChange={handleSelectAll}
                            className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500"
                          />
                        </th>
                      )}
                      <th 
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                        onClick={() => handleSort('name')}
                      >
                        <div className="flex items-center space-x-1">
                          <span>Name</span>
                          {sortField === 'name' && (
                            sortOrder === 'asc' ? <SortAsc className="h-3 w-3" /> : <SortDesc className="h-3 w-3" />
                          )}
                        </div>
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Contact
                      </th>
                      <th 
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                        onClick={() => handleSort('employer')}
                      >
                        <div className="flex items-center space-x-1">
                          <span>Employer</span>
                          {sortField === 'employer' && (
                            sortOrder === 'asc' ? <SortAsc className="h-3 w-3" /> : <SortDesc className="h-3 w-3" />
                          )}
                        </div>
                      </th>
                      <th 
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                        onClick={() => handleSort('agency')}
                      >
                        <div className="flex items-center space-x-1">
                          <span>Agency</span>
                          {sortField === 'agency' && (
                            sortOrder === 'asc' ? <SortAsc className="h-3 w-3" /> : <SortDesc className="h-3 w-3" />
                          )}
                        </div>
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Employment
                      </th>
                      <th 
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                        onClick={() => handleSort('status')}
                      >
                        <div className="flex items-center space-x-1">
                          <span>Status</span>
                          {sortField === 'status' && (
                            sortOrder === 'asc' ? <SortAsc className="h-3 w-3" /> : <SortDesc className="h-3 w-3" />
                          )}
                        </div>
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Documents
                      </th>
                      <th 
                        className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer hover:bg-gray-100"
                        onClick={() => handleSort('updatedAt')}
                      >
                        <div className="flex items-center space-x-1">
                          <span>Updated</span>
                          {sortField === 'updatedAt' && (
                            sortOrder === 'asc' ? <SortAsc className="h-3 w-3" /> : <SortDesc className="h-3 w-3" />
                          )}
                        </div>
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Actions
                      </th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {filteredAndSortedHousemaids.map((housemaid) => (
                      <HousemaidRow key={housemaid.id} housemaid={housemaid} />
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}

      {/* Import Modal */}
      {showImport && (
        <ExcelImport
          onImport={onBulkImport}
          onClose={() => setShowImport(false)}
        />
      )}

      {/* Reports Modal */}
      {showReports && (
        <ReportGenerator
          housemaids={filteredAndSortedHousemaids}
          onClose={() => setShowReports(false)}
        />
      )}
    </div>
  );
};

export default HousemaidList;