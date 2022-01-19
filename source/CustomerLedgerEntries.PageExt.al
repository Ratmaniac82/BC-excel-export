pageextension 50110 "PPI Customer Ledger Entries" extends "Customer Ledger Entries"
{
    actions
    {
        addafter("F&unctions")
        {
            action(NECExportToExcel)
            {
                Caption = 'Export to Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Export;

                trigger OnAction()
                var
                begin
                    ExportCustLedgerEntries(Rec);
                end;
            }
        }
    }

    local procedure ExportCustLedgerEntries(var CustLedgerEntryRec: Record "Cust. Ledger Entry")
    var
        TempExcelBuffer: Record "Excel Buffer" temporary; // Main method
        CustLedgerEntriesLbl: Label 'Customer Ledger Entries'; // Sheet Name
        ExcelFileName: Label 'CustomerLedgerEntries_%1_%2'; // Excel File Name
    begin
        TempExcelBuffer.Reset(); // Clear Excel Buffer
        TempExcelBuffer.DeleteAll(); // Clear Excel Buffer
        TempExcelBuffer.NewRow(); // Start adding new row in Excel // Header of content
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Entry No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text); // Start adding new Column
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Posting Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Document Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Document No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Customer No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Customer Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Currency Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Original Amount"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption(Amount), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Amount (LCY)"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Remaining Amount"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption("Remaining Amt. (LCY)"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(CustLedgerEntryRec.FieldCaption(Open), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        if CustLedgerEntryRec.FindSet() then
            repeat // Add Customer Ledger Entry lines to excel
                TempExcelBuffer.NewRow();
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Entry No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Posting Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Document Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Document No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Customer No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Customer Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec.Description, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Currency Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Original Amount", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec.Amount, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Amount (LCY)", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Remaining Amount", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec."Remaining Amt. (LCY)", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(CustLedgerEntryRec.Open, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
            until CustLedgerEntryRec.Next() = 0;
        // Create a excel and save
        TempExcelBuffer.CreateNewBook(CustLedgerEntriesLbl);
        TempExcelBuffer.WriteSheet(CustLedgerEntriesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileName, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
    end;
}