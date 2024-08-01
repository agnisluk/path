codeunit 50000 "Sharepoint Management SP"
{
    procedure SaveFile(FileDirectory: Text; FileName: Text; IS: InStream): Text
    var
        TempSharePointFile: Record "SharePoint File" temporary;
        Diag: Interface "HTTP Diagnostics";
    begin
        InitializeConnection(); //Initialize a connection if not connected
        if SharePointClient.AddFileToFolder(FileDirectory, FileName, IS, TempSharePointFile) then //Use AddFileToFolder to create a file from a stream
            exit(TempSharePointFile."Server Relative Url")
        else begin
            Diag := SharePointClient.GetDiagnostics();
            if (not Diag.IsSuccessStatusCode()) then
                Error(DiagErr, Diag.GetErrorMessage());
        end;
    end;

    procedure OpenFile(FileDirectory: Text)
    var
        TempSharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileNotFoundErr: Label 'File not found.';
        FileName: Text;
    begin
        InitializeConnection();

        FileName := FileMgt.GetFileName(FileDirectory);
        if FileName = '' then
            Error(FileNotFoundErr);

        FileDirectory := FileMgt.GetDirectoryName(FileDirectory);
        if not SharePointClient.GetFolderFilesByServerRelativeUrl(FileDirectory, TempSharePointFile) then
            Error(FileNotFoundErr);

        TempSharePointFile.SetRange(Name, FileName);
        if TempSharePointFile.FindFirst() then
            SharePointClient.DownloadFileContent(TempSharePointFile.OdataId, FileName)
        else
            Error(FileNotFoundErr);
    end;

    local procedure InitializeConnection()
    var
        SharepointSetup: Record "Sharepoint Connector Setup SP";
        TempSharePointList: Record "SharePoint List" temporary;
        Diag: Interface "HTTP Diagnostics";
        AadTenantId: Text;
    begin
        if Connected then
            exit;

        SharepointSetup.Get();

        AadTenantId := GetAadTenantNameFromBaseUrl(SharepointSetup."Sharepoint URL"); //Used to get an Azure Active Directory ID from a URL
        SharePointClient.Initialize(SharepointSetup."Sharepoint URL", GetSharePointAuthorization(AadTenantId));

        SharePointClient.GetLists(TempSharePointList);

        Diag := SharePointClient.GetDiagnostics();

        if (not Diag.IsSuccessStatusCode()) then
            Error(DiagErr, Diag.GetErrorMessage());

        Connected := true;
    end;

    local procedure GetSharePointAuthorization(AadTenantId: Text): Interface "SharePoint Authorization"
    var
        SharepointSetup: Record "Sharepoint Connector Setup SP";
        SharePointAuth: Codeunit "SharePoint Auth.";
        Scopes: List of [Text];
    begin
        SharepointSetup.Get();

        Scopes.Add('00000003-0000-0ff1-ce00-000000000000/.default'); //Using a default scope provided as an example
                                                                     //CAN put field in Setup
        SetPassword(SharepointSetup."Client ID");
        exit(SharePointAuth.CreateAuthorizationCode(AadTenantId, SharepointSetup."Client ID"/*SecretText*/, SharepointSetup."Client Secret", Scopes));
    end;

    local procedure GetAadTenantNameFromBaseUrl(BaseUrl: Text): Text
    var
        Uri: Codeunit Uri;
        MySiteHostSuffixTxt: Label '-my.sharepoint.com', Locked = true;
        OnMicrosoftTxt: Label '.onmicrosoft.com', Locked = true;
        SharePointHostSuffixTxt: Label '.sharepoint.com', Locked = true;
        UrlInvalidErr: Label 'The Base Url %1 does not seem to be a valid SharePoint Online Url.', Comment = '%1=BaseUrl';
        Host: Text;
    begin
        Uri.Init(BaseUrl);
        Host := Uri.GetHost();
        if not Host.EndsWith(SharePointHostSuffixTxt) then
            Error(UrlInvalidErr, BaseUrl);
        if Host.EndsWith(MySiteHostSuffixTxt) then
            exit(CopyStr(Host, 1, StrPos(Host, MySiteHostSuffixTxt) - 1) + OnMicrosoftTxt);
        exit(CopyStr(Host, 1, StrPos(Host, SharePointHostSuffixTxt) - 1) + OnMicrosoftTxt);
    end;

    [NonDebuggable]
    local procedure SetPassword(SetSecretText: Text[250])
    begin
        SecretText := SetSecretText;
    end;

    var
        SharePointClient: Codeunit "SharePoint Client";
        SecretText: SecretText;
        Connected: Boolean;
        DiagErr: Label 'Sharepoint Management error:\\%1', Comment = '%1 = GetErrorMessage';
}