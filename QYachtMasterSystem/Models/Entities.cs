using System;

namespace QYachtMaster.Models
{
    public class Customer
    {
        public int CustomerID { get; set; }
        public string Name { get; set; } = string.Empty;
        public string MobilePhone { get; set; } = string.Empty;
        public string? WorkPhone { get; set; }
        public string? OtherPhone { get; set; }
        public string? Address { get; set; }
        public double TotalDebt { get; set; } = 0.0;
        public string CreatedAt { get; set; } = string.Empty;
        public string? UpdatedAt { get; set; }
    }

    public class Vehicle
    {
        public int VehicleID { get; set; }
        public int CustomerID { get; set; }
        public string? RegNo { get; set; }
        public string? Year { get; set; }
        public string? Make { get; set; }
        public string? Model { get; set; }
        public string? Color { get; set; }
        public string? LastOdometer { get; set; }
    }

    public class JobCard
    {
        public int JobCardID { get; set; }
        public string JobCardNo { get; set; } = string.Empty; // Format: JC-2026-XXXX
        public int CustomerID { get; set; }
        public int VehicleID { get; set; }
        public string? ArrivalTime { get; set; }
        public string JobDate { get; set; } = string.Empty;
        public string? PickupTime { get; set; }
        public string? PickupPeriod { get; set; } // 'AM' or 'PM'
        public string? DropoffTime { get; set; }
        public string? DropoffPeriod { get; set; } // 'AM' or 'PM'
        public string? OdometerAtVisit { get; set; }
        public string Status { get; set; } = "In Progress"; // In Progress | Completed | Cancelled
        public string CreatedAt { get; set; } = string.Empty;
        public string? CreatedBy { get; set; }
    }

    public class JobCardRepairLine
    {
        public int LineID { get; set; }
        public int JobCardID { get; set; }
        public int LineNumber { get; set; } // Range: 1 to 8
        public string? ServiceDescription { get; set; }
        public string? TechnicianName { get; set; }
    }

    public class JobCardTechComment
    {
        public int CommentID { get; set; }
        public int JobCardID { get; set; }
        public string ServiceKey { get; set; } = string.Empty;
        public int IsSelected { get; set; } = 0; // 0 = false, 1 = true
        public string? Notes { get; set; }
    }

    public class Invoice
    {
        public int InvoiceID { get; set; }
        public string InvoiceNo { get; set; } = string.Empty; // Format: INV-2026-XXXX (starting 1000, 1002...)
        public int JobCardID { get; set; }
        public int CustomerID { get; set; }
        public int IsLocked { get; set; } = 0;
        public string? LockedAt { get; set; }
        public string? LockedBy { get; set; }
        public double SubTotal { get; set; } = 0.0;
        public double LaborTotal { get; set; } = 0.0;
        public double GrandTotal { get; set; } = 0.0;
        public double PaidAmount { get; set; } = 0.0;
        public double RemainingDebt { get; set; } = 0.0;
        public string? PaymentMethod { get; set; } // Cash | BankTransfer | Partial
        public string CreatedAt { get; set; } = string.Empty;
    }

    public class InvoiceLine
    {
        public int LineID { get; set; }
        public int InvoiceID { get; set; }
        public string Description { get; set; } = string.Empty;
        public double Qty { get; set; } = 1.0;
        public double UnitPrice { get; set; } = 0.0;
        public double UnitCostPrice { get; set; } = 0.0;
        public double QtySubTotal => Qty * UnitPrice; // Virtual computed value
        public double DiscountAmount { get; set; } = 0.0;
        public string? DiscountNote { get; set; }
        public double FinalLineTotal { get; set; } = 0.0;
        public int IsManualOverride { get; set; } = 0; // 0 = false, 1 = true
        public string? SourceServiceKey { get; set; }
        public int SortOrder { get; set; } = 0;
    }

    public class Part
    {
        public int PartID { get; set; }
        public string PartNo { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string? Category { get; set; }
        public string? Prefix { get; set; }
        public double QtyInStock { get; set; } = 0.0;
        public double CostPrice { get; set; } = 0.0;
        public double SalePrice { get; set; } = 0.0;
        public string? Location { get; set; }
        public string? CarBrand { get; set; }
        public string? CarModel { get; set; }
        public string? LastSyncedAt { get; set; }
        public string? LastUpdatedAt { get; set; }
    }

    public class Payment
    {
        public int PaymentID { get; set; }
        public int CustomerID { get; set; }
        public int? InvoiceID { get; set; }
        public double Amount { get; set; }
        public string PaymentType { get; set; } = "Cash"; // Cash | BankTransfer
        public string? ReceivedBy { get; set; }
        public string ReceivedAt { get; set; } = string.Empty;
        public string? Notes { get; set; }
    }

    public class AuditTrail
    {
        public int AuditID { get; set; }
        public string ActionType { get; set; } = string.Empty; // DEBT_CLEARED | INVOICE_LOCKED | FOLDER_ACCESS | FAILED_ACCESS | CASH_RECEIVED
        public string? EntityType { get; set; }
        public int? EntityID { get; set; }
        public string? Description { get; set; }
        public string? OldValue { get; set; }
        public string? NewValue { get; set; }
        public string? UserName { get; set; }
        public string? IPAddress { get; set; }
        public string Timestamp { get; set; } = string.Empty; // ISO 8601 string
    }

    public class User
    {
        public int UserID { get; set; }
        public string Username { get; set; } = string.Empty;
        public string PasswordHash { get; set; } = string.Empty;
        public string Role { get; set; } = "Employee"; // Admin | Employee
        public int IsActive { get; set; } = 1;
    }

    public class PartReturn
    {
        public int ReturnID { get; set; }
        public int? InvoiceID { get; set; }           // Nullable — employee may not know exact InvoiceID
        public string InvoiceNo { get; set; } = string.Empty;  // Human-readable reference
        public int CustomerID { get; set; }
        public int? PartID { get; set; }              // Resolved from PartNo lookup
        public string PartNo { get; set; } = string.Empty;
        public double QtyReturned { get; set; } = 1.0;
        public double UnitPrice { get; set; } = 0.0;
        public double RefundAmount { get; set; } = 0.0;
        /// <summary>Cash = cash refund payment (negative payment); DebtReduction = subtract from customer TotalDebt</summary>
        public string RefundType { get; set; } = "Cash"; // Cash | DebtReduction
        public string? ProcessedBy { get; set; }
        public string ProcessedAt { get; set; } = string.Empty;
        public string? Notes { get; set; }
    }
}
