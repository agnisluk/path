page 50000 "Sharepoint Connector Setup SP"
{
    ApplicationArea = Basic, Suite;
    Caption = 'Sharepoint Connector Setup';
    DeleteAllowed = false;
    InsertAllowed = false;
    PageType = Card;
    SourceTable = "Sharepoint Connector Setup SP";
    UsageCategory = Administration;

    layout
    {
        area(Content)
        {
            group(Setup)
            {
                Caption = 'Sharepoint Setup';
                field("Client ID"; Rec."Client ID")
                {
                    ApplicationArea = All;
                    Caption = 'Client ID';
                    ToolTip = 'Specifies the value of the Client ID field.';
                    Editable = Rec."Use Sharepoint";
                }
                field("Client Secret"; Rec."Client Secret")
                {
                    ApplicationArea = All;
                    ExtendedDatatype = Masked;
                    Caption = 'Client Secret';
                    ToolTip = 'Specifies the value of the Client Secret field.';
                    Editable = Rec."Use Sharepoint";
                }
                field("Sharepoint URL"; Rec."Sharepoint URL")
                {
                    ApplicationArea = All;
                    Caption = 'Sharepoint URL';
                    ToolTip = 'Specifies the value of the Sharepoint URL field.';
                    Editable = Rec."Use Sharepoint";
                }
                field("Parent Folder URL"; Rec."Parent Folder URL")
                {
                    ApplicationArea = All;
                    Caption = 'Parent Folder URL';
                    ToolTip = 'Specifies the value of the Parent Folder URL field.';
                    Editable = Rec."Use Sharepoint";
                }
                field("Use Sharepoint"; Rec."Use Sharepoint")
                {
                    ApplicationArea = All;
                    Caption = 'Use Sharepoint';
                    ToolTip = 'Specifies the value of the Use Sharepoint field.';
                }
            }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.Reset();
        if not Rec.Get() then begin
            Rec.Init();
            Rec.Insert();
        end
    end;
}