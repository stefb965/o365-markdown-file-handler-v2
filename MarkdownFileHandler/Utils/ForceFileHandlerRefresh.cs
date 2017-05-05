/*
 * Markdown File Handler - Sample Code
 * Copyright (c) Microsoft Corporation
 * All rights reserved. 
 * 
 * MIT License
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of 
 * this software and associated documentation files (the ""Software""), to deal in 
 * the Software without restriction, including without limitation the rights to use, 
 * copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
 * Software, and to permit persons to whom the Software is furnished to do so, 
 * subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all 
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
 * PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
 * HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE 
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */


namespace MarkdownFileHandler
{
    using FileHandlerActions;
    using System;
    using System.Threading.Tasks;

    public class ForceFileHandlerRefresh
    {
        /// <summary>
        /// Force the refresh for file handlers cached in SharePoint / OneDrive
        /// </summary>
        public static async Task ForceFileHandlerRefreshAsync(string oneDriveUrl)
        {
            try
            {
                var uri = new Uri(oneDriveUrl);
                var resourceId = uri.GetLeftPart(UriPartial.Authority);
                var accessToken = await AuthHelper.GetUserAccessTokenSilentAsync(resourceId);

                if (accessToken != null)
                {
                    // Call SharePoint and force file handlers to refresh
                    UriBuilder builder = new UriBuilder(uri);
                    builder.Path = "/_api/v2.0/drive/apps";
                    builder.Query = "forceRefresh=1";
                    await HttpHelper.Default.GetMetadataForUrlAsync<object>(builder.ToString(), accessToken);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to force refresh of file handlers. {ex.Message}");
            }
        }
    }
}