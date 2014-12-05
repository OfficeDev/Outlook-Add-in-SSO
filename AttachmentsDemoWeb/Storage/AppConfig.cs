namespace AttachmentsDemoWeb.Storage
{
    // Represents an entry in the app config cache
    public class AppConfig
    {
        // The OAuth refresh token for the user
        public string RefreshToken { get; set; }
        // The user's OneDrive resource id
        public string OneDriveResourceId { get; set; }
        // The user's OneDrive API endpoint URL
        public string OneDriveApiEndpoint { get; set; }
    }
}