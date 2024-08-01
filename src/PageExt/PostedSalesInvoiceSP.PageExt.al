pageextension 50000 "Posted Sales Invoice SP" extends "Posted Sales Invoice"
{
    layout
    {
        addfirst(FactBoxes)
        {
            systempart(PaymentTermsLinks; Links)
            {
                ApplicationArea = RecordLinks;
                Caption = 'Sharepoint Links';
            }
        }
    }
}
