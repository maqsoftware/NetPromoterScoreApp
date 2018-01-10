using Amazon;
using Amazon.Rekognition;
using Amazon.Rekognition.Model;
using Amazon.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Security.Credentials;

namespace EmotionRecognitionAPI
{
    public sealed class EmotionDetectionHandler
    {

        public static IAsyncOperation<string> GetFaceDetails(string base64, string AccessKey, string SecretKey)
        {
            return Task.Run<string>(async () =>
            {
                byte[] imageBytes;
                try
                {
                    base64 = base64.Substring(base64.IndexOf(',') + 1).Trim('\0');                    
                    imageBytes = System.Convert.FromBase64String(base64);
                }
                catch (Exception e) {
                    return e.Message;
                }                
                string sJSONResponse = "";                
                AWSCredentials credentials;
                try
                {                    
                    credentials = new BasicAWSCredentials(AccessKey, SecretKey);                    
                }
                catch (Exception e)
                {
                    throw new AmazonClientException("Cannot load the credentials from the credential profiles file. "
                            + "Please make sure that your credentials file is at the correct "
                            + "location (/Users/<userid>/.aws/credentials), and is in a valid format.", e);
                }
                DetectFacesRequest request = new DetectFacesRequest { Attributes = new List<string>(new string[] { "ALL" }) };
                DetectFacesResponse result = null;
                request.Image = new Image { Bytes = new MemoryStream(imageBytes, 0, imageBytes.Length) };
                AmazonRekognitionClient rekognitionClient = new AmazonRekognitionClient(credentials, RegionEndpoint.USWest2);
                try
                {
                    result = await rekognitionClient.DetectFacesAsync(request);
                }
                catch (AmazonRekognitionException e)
                {
                    throw e;
                }
                // Return server status as unhealthy with appropriate status code
                sJSONResponse = JsonConvert.SerializeObject(result.FaceDetails);
                return sJSONResponse;
            }).AsAsyncOperation();
            
        }             
    }  
}
