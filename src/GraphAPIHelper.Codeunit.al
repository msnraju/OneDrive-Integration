codeunit 50115 "Graph API Helper"
{
    var
        OAuth2: Codeunit OAuth2;
        ClientIdTxt: Label '96e2efa1-a6fb-4f04-97d9-1f9ac9c15917', Locked = true;
        ClientSecret: Label '44Qi99bD4EC4S27~_.5htAp1o_lLd7tfBg', Locked = true;
        ResourceUrlTxt: Label 'https://graph.microsoft.com', Locked = true;
        OAuthAuthorityUrlTxt: Label 'https://login.microsoftonline.com/67c5a58a-7424-4d4d-b6c2-ddc89830cf74/oauth2/authorize', Locked = true;
        RedirectURLTxt: Label 'http://localhost:8080/BC160/OAuthLanding.htm', Locked = true;
        OneDriveRootQueryUri: Label 'https://graph.microsoft.com/v1.0/me/drive/root/children', Locked = true;

    procedure GetAccessToken(): Text
    var
        PromptInteraction: Enum "Prompt Interaction";
        AccessToken: Text;
        AuthCodeError: Text;
    begin
        OAuth2.AcquireTokenByAuthorizationCode(
            ClientIdTxt,
            ClientSecret,
            OAuthAuthorityUrlTxt,
            RedirectURLTxt,
            ResourceURLTxt,
            PromptInteraction::Consent,
            AccessToken,
            AuthCodeError);

        if (AccessToken = '') or (AuthCodeError <> '') then
            Error(AuthCodeError);

        exit(AccessToken);
    end;

    procedure GetOneDriveFiles(): JsonObject
    var
        Client: HttpClient;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        JsonResponse: JsonObject;
        AccessToken: Text;
        JsonContent: Text;
    begin
        AccessToken := GetAccessToken();

        RequestMessage.Method('GET');
        RequestMessage.SetRequestUri(OneDriveRootQueryUri);
        Client.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));
        Client.DefaultRequestHeaders().Add('Accept', 'application/json');

        if Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.HttpStatusCode() = 200 then begin
                ResponseMessage.Content.ReadAs(JsonContent);
                JsonResponse.ReadFrom(JsonContent);
                exit(JsonResponse);
            end;
    end;
}