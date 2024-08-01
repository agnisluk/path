table 50000 "Sharepoint Connector Setup SP"
{
    DataClassification = ToBeClassified;

    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            DataClassification = CustomerContent;
            NotBlank = true;
            Caption = 'Primary Key';
        }
        field(2; "Client ID"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
        }
        field(3; "Client Secret"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
        }
        field(4; "Sharepoint URL"; Text[250])
        {
            DataClassification = CustomerContent;
            Caption = 'Sharepoint URL';
        }
        field(5; "Parent Folder URL"; Text[250])
        {
            DataClassification = CustomerContent;
            Caption = 'Parent Folder URL';
        }
        field(6; "Use Sharepoint"; Boolean)
        {
            DataClassification = CustomerContent;
            Caption = 'Use Sharepoint';
        }
    }

    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }
}