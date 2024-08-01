codeunit 50001 Events
{
    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Document-Mailing", OnAfterEmailSentSuccesfully, '', false, false)]
    local procedure "Document-Mailing_OnAfterEmailSentSuccesfully"(var TempEmailItem: Record "Email Item" temporary; PostedDocNo: Code[20]; ReportUsage: Integer)
    var
        SalesInvoiceHeader: Record "Sales Invoice Header";
        SharepointConnectorSetup: Record "Sharepoint Connector Setup SP";
        SharepointMgt: Codeunit "Sharepoint Management SP";
        Attachment: Codeunit "Temp Blob";
        Attachments: Codeunit "Temp Blob List";
        AttachmentStream: Instream;
        Index: Integer;
        AttachmentNames: List of [Text];
        AttachFileName: Text;
        LinkDescription: Text;
        LinkUrl: Text;
        ServerRelativeUrl: Text;
    begin
        if not SharepointConnectorSetup.Get() then
            exit;
        if not SharepointConnectorSetup."Use Sharepoint" then
            exit;

        TempEmailItem.GetAttachments(Attachments, AttachmentNames);
        for Index := 1 to Attachments.Count() do begin
            Attachments.Get(Index, Attachment);
            Attachment.CreateInStream(AttachmentStream);

            AttachFileName := GetAttachmentName(AttachmentNames, Index);
            LinkDescription := 'Open ' + AttachFileName + ' in Sharepoint ';
            LinkUrl := SharepointConnectorSetup."Sharepoint URL" + SharepointConnectorSetup."Parent Folder URL" + '/' + AttachFileName;
            ServerRelativeUrl := SharepointMgt.SaveFile(SharepointConnectorSetup."Parent Folder URL", AttachFileName, AttachmentStream);

            if ServerRelativeUrl <> '' then begin
                if not SalesInvoiceHeader.Get(PostedDocNo) then
                    exit;
                SalesInvoiceHeader.AddLink(LinkUrl, LinkDescription)
            end;
        end;
    end;

    local procedure GetAttachmentName(AttachmentNames: List of [Text]; Index: Integer) AttachmentName: Text[250]
    begin
        AttachmentName := CopyStr(AttachmentNames.Get(Index), 1, 250);
    end;


}
