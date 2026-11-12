export type UserRole = 'SERVICE TEAM' | 'TECHNICIAN' | 'SUPERVISOR' | null;

export type TicketStatus =
  | 'NEW'
  | 'ASSIGNED'
  | 'INSPECTION'
  | 'REPAIRING'
  | 'PENDING_PARTS_APPROVAL'
  | 'PENDING_SUPERVISOR_APPROVAL'
  | 'REJECTED'
  | 'CUSTOMER_ACCEPTANCE'
  | 'CLOSE';

export interface ChecklistItem {
  item: string;
  measuredValue?: string;
  unit?: string;
  acceptanceCriteria?: string;
  beforeStatus: 'OK' | 'NG' | 'N/A';
  afterStatus: 'OK' | 'NG' | 'N/A';
}

export interface PartItem {
  id?: string;
  name: string;
  qty: number;
  unit: string;
  price: number;
  total: number;
  billingType?: 'WARRANTY_INTERNAL' | 'CUSTOMER_CHARGE';
}

export type RepairStatus = TicketStatus;
export interface LineNotifyLog {
  id: string;
  ticketId: string;
  message: string;
  timestamp: string;
  status: 'SENT' | 'FAILED';
}

export interface EquipmentItem {
  type: string;
  brand: string;
  model: string;
  tag?: string;
  symptom: string;
  detail?: string;
  beforePhotos: string[];
  afterPhotos: string[];
  analysis: string;
  cause: string;
  standardsReference?: string;
  faultOrigin: string;
  correctiveAction: string;
  parts: PartItem[];
  partsApproved: boolean;
  checklistResults: ChecklistItem[];
}

export interface ServiceTicket {
  id: string;
  company: string;
  taxId: string;
  docNo: string;
  contact: string;
  siteContact: string;
  location: string;
  warranty: string;
  serviceType: string;
  urgency: string;
  desiredDate: string;
  confirmedDate: string;
  status: TicketStatus;
  isLocked: boolean;
  rejectReason: string;
  createdAt: string;
  equipmentList: EquipmentItem[];
  assignedTech: string;
  leadTech?: string;
  approvedBySupervisor: string;
  customerSignature?: string;
  customerSignerName?: string;
  customerSignedAt?: string;
  technicianSignature?: string;
  technicianSignedAt?: string;
}

export interface Technician {
  id: string;
  name: string;
  position?: string;
  specialization: string;
  employmentStatus?: string;
  status: 'Available' | 'On Duty';
  pin: string;
}

export interface TroubleshootingTemplate {
  id: string;
  title: string;
  category: string;
  sourceRef: string;
  analysis: string;
  cause: string;
  correctiveAction: string;
  recommendedChecklist?: string[];
}

export type RepairTicket = ServiceTicket;
