using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Timers;
using System.IO;

//ML References

using System.Drawing;
using System.Drawing.Drawing2D;
using ReadRPIHubEmails.YoloParser;
using ReadRPIHubEmails.DataStructures;
using ReadRPIHubEmails;
using Microsoft.ML;

namespace ReadRPIHubEmails
{
    class Program
    {
        private static System.Timers.Timer cTimer;
        private static bool Working;

        static void Main(string[] args)
        {
            cTimer = new Timer();
            cTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            //cTimer.Elapsed += new ElapsedEventHandler(FoundPerson);
            cTimer.Interval = 60000;
            cTimer.Enabled = true;
            Working = false;
            Console.Read();
        }
        private static void OnSyncEnd()
        {
            Working = false;
        }
        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application oApp;
            Microsoft.Office.Interop.Outlook.NameSpace oNs;
            Microsoft.Office.Interop.Outlook.Stores oStores;
            Microsoft.Office.Interop.Outlook.Store oStore;
            Microsoft.Office.Interop.Outlook.MAPIFolder oFolder;
            Microsoft.Office.Interop.Outlook.MAPIFolder oReadFolder;
            Microsoft.Office.Interop.Outlook.MAPIFolder oDeleteFolder;
            Microsoft.Office.Interop.Outlook.MAPIFolder oFolderStore;
            Microsoft.Office.Interop.Outlook.Items oItems;
            //Microsoft.Office.Interop.Outlook.MailItem oMailItem;
            Microsoft.Office.Interop.Outlook.SyncObjects oSyncObjects;
            Microsoft.Office.Interop.Outlook.SyncObject oSyncObject;
            int ReadCnt = 0;
            int DelCnt = 0;
            var assetsRelativePath = @"../../../assets";
            string assetsPath = GetAbsolutePath(assetsRelativePath);
            var modelFilePath = Path.Combine(assetsPath, "Model", "TinyYolo2_model.onnx");
            var imagesFolder = Path.Combine(assetsPath, "images");
            var outputFolder = Path.Combine(assetsPath, "images", "output");
            var savedFolder = Path.Combine(assetsPath, "images", "saved");
            var attFolder = Path.Combine(assetsPath, "images", "emailfolders");
            cTimer.Stop();
            // Initialize MLContext
            MLContext mlContext = new MLContext();
            //Check if syncing if so quit
            if (Working)
            {
                return;
            }
            try
            {
                oApp = (Microsoft.Office.Interop.Outlook.Application)Microsoft.VisualBasic.Interaction.GetObject(null, "Outlook.Application");
                oNs = oApp.GetNamespace("MAPI");
                oStores = oNs.Stores;
                oStore = null;
                for (int i = 1; i < oStores.Count + 1; i++)
                {
                    if (oStores[i].DisplayName.Contains("Norrisrpihub@hotmail.com"))
                    {
                        oStore = oStores[i];
                    }
                }
                if (oStore == null)
                {
                    throw new Exception("Could not find Norrisrpihub@hotmail.com");
                }
                oFolderStore = oStore.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox); //Inbox
                oFolder = oNs.Folders["Norrisrpihub@hotmail.com"].Folders["Inbox"];
                oReadFolder = oFolder.Folders["Read"];
                oDeleteFolder = oFolder.Folders["Delete"];
                oItems = oFolder.Items.Restrict("[Unread] = true");
                oItems = oFolder.Items;
                for (int i = 1; i < oItems.Count; i++)
                {
                    if (oItems[i].Class == 43)
                    {

                        //oMailItem = oItems[i];
                        string nFolder = attFolder + @"\" + oItems[i].ReceivedTime.ToString("yyyy-MM-dd-HH-mm-ss");
                        Directory.CreateDirectory(nFolder);

                        var emailFolder = Path.GetFullPath(nFolder);
                        oItems[i].UnRead = false;
                        ReadCnt++;
                        long totSize = 0;
                        //Delete it if there's no attachments
                        if (oItems[i].Attachments.Count == 0)
                        {
                            oItems[i].Move(oDeleteFolder);                          
                            DelCnt++;
                            
                        }
                        else
                        {
                            for(int j = 1; j< oItems[i].Attachments.Count; j++)
                            {
                                totSize += oItems[i].Attachments[j].Size;
                                if (oItems[i].Attachments[j].Size > 1000)
                                {
                                    if (oItems[i].Attachments[j].FileName.Substring(oItems[i].Attachments[j].FileName.Length-4,4) == ".jpg")
                                    {
                                        //Save file to email folder for processing
                                        Console.WriteLine(emailFolder + @"\" + oItems[i].ReceivedTime.ToString("yyyy-dd-M-HH-mm-ss") + " " + j + ".jpg");
                                        oItems[i].Attachments[j].SaveAsFile(emailFolder + @"\" +  oItems[i].ReceivedTime.ToString("yyyy-dd-M-HH-mm-ss") + " " + j + ".jpg");
                                    }
                                    
                                }
                            }
                            if (totSize < 1000)
                            {
                                //Couldn't be a reasonably sized jpeg, so delete
                                oItems[i].Move(oDeleteFolder);
                                DelCnt++;

                            }
                            else if (!FoundPerson(emailFolder.ToString()))
                            {
                                oItems[i].Move(oDeleteFolder);
                                DelCnt++;

                            }
                            else
                            {
                                //copy all attachments to the saved folder
                                System.IO.DirectoryInfo di = new DirectoryInfo(emailFolder);
                                foreach (FileInfo file in di.GetFiles())
                                {
                                    if (file.Name.Contains(".jpg"))
                                    {
                                        file.CopyTo(savedFolder + @"\" + file.Name,true);
                                    }
                                }
                                oItems[i].Move(oReadFolder);
                            }
                        }
                        Console.WriteLine("Processed: " + ReadCnt.ToString() + " Emails.");
                    }
                }
                //Sync to server
                Console.WriteLine("Syncing to Server...");
                oSyncObjects = oNs.SyncObjects;
                oSyncObject = oSyncObjects["All Accounts"];
                oSyncObject.SyncEnd += new Microsoft.Office.Interop.Outlook.SyncObjectEvents_SyncEndEventHandler(OnSyncEnd);
                Working = true;
                oSyncObject.Start();

                oApp = null;
                oSyncObject = null;
                Console.WriteLine("Completed Processing at " + DateTime.Now.ToString(@"MM/dd/yyyy HH:mm:ss") + " DELETED: " + DelCnt + "; READ: " + ReadCnt);
                cTimer.Start();
            }

            catch (Exception ex)
            {
                oSyncObject = null;
                oSyncObjects = null;
                oApp = null;
                oNs = null;
                oItems = null;
                oStores = null;
                oStore = null;
                Console.WriteLine("Error:" + ex.Message);
                Working = false;
                cTimer.Enabled = true;
            }

        }
        private static bool FoundPerson(string inEmailFolder)
        {
            var assetsRelativePath = @"../../../assets";
            string assetsPath = GetAbsolutePath(assetsRelativePath);
            var modelFilePath = Path.Combine(assetsPath, "Model", "TinyYolo2_model.onnx");
            //var imagesFolder = Path.Combine(assetsPath, "images");
            var imagesFolder = Path.GetFullPath(inEmailFolder);
            var outputFolder = Path.Combine(assetsPath, "images", "output");
            long personCnt = 0;

            // Initialize MLContext
            MLContext mlContext = new MLContext();
            cTimer.Stop();
            try
            {
                // Load Data
                IEnumerable<ImageNetData> images = ImageNetData.ReadFromFile(imagesFolder);
                IDataView imageDataView = mlContext.Data.LoadFromEnumerable(images);

                // Create instance of model scorer
                var modelScorer = new OnnxModelScorer(imagesFolder, modelFilePath, mlContext);

                // Use model to score data
                IEnumerable<float[]> probabilities = modelScorer.Score(imageDataView);

                // Post-process model output
                YoloOutputParser parser = new YoloOutputParser();

                var boundingBoxes =
                    probabilities
                    .Select(probability => parser.ParseOutputs(probability))
                    .Select(boxes => parser.FilterBoundingBoxes(boxes, 5, .5F));

                // Draw bounding boxes for detected objects in each of the images
                for (var i = 0; i < images.Count(); i++)
                {
                    string imageFileName = images.ElementAt(i).Label;
                    IList<YoloBoundingBox> detectedObjects = boundingBoxes.ElementAt(i);
                    //This is added to specifically count the objects that are people. Because this is for a security camera I only really care if a person is found, however there are many other objects that would be included as well.
                    for(var j = 0; j < detectedObjects.Count(); j++)
                    {
                        if (detectedObjects[j].Label == "person")
                        {
                            personCnt++;
                        }
                    }
                    DrawBoundingBox(imagesFolder, outputFolder, imageFileName, detectedObjects);

                    LogDetectedObjects(imageFileName, detectedObjects);
                }
                if(personCnt == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                cTimer.Start();
                Console.WriteLine(ex.ToString());
                return false;
            }
        }
        private static void DrawBoundingBox(string inputImageLocation, string outputImageLocation, string imageName, IList<YoloBoundingBox> filteredBoundingBoxes)
        {
            Image image = Image.FromFile(Path.Combine(inputImageLocation, imageName));

            var originalImageHeight = image.Height;
            var originalImageWidth = image.Width;
            if (filteredBoundingBoxes.Count > 0)
            {
                foreach (var box in filteredBoundingBoxes)
                {
                    // Get Bounding Box Dimensions
                    var x = (uint)Math.Max(box.Dimensions.X, 0);
                    var y = (uint)Math.Max(box.Dimensions.Y, 0);
                    var width = (uint)Math.Min(originalImageWidth - x, box.Dimensions.Width);
                    var height = (uint)Math.Min(originalImageHeight - y, box.Dimensions.Height);

                    // Resize To Image
                    x = (uint)originalImageWidth * x / OnnxModelScorer.ImageNetSettings.imageWidth;
                    y = (uint)originalImageHeight * y / OnnxModelScorer.ImageNetSettings.imageHeight;
                    width = (uint)originalImageWidth * width / OnnxModelScorer.ImageNetSettings.imageWidth;
                    height = (uint)originalImageHeight * height / OnnxModelScorer.ImageNetSettings.imageHeight;

                    // Bounding Box Text
                    string text = $"{box.Label} ({(box.Confidence * 100).ToString("0")}%)";

                    using (Graphics thumbnailGraphic = Graphics.FromImage(image))
                    {
                        thumbnailGraphic.CompositingQuality = CompositingQuality.HighQuality;
                        thumbnailGraphic.SmoothingMode = SmoothingMode.HighQuality;
                        thumbnailGraphic.InterpolationMode = InterpolationMode.HighQualityBicubic;

                        // Define Text Options
                        Font drawFont = new Font("Arial", 12, FontStyle.Bold);
                        SizeF size = thumbnailGraphic.MeasureString(text, drawFont);
                        SolidBrush fontBrush = new SolidBrush(Color.Black);
                        Point atPoint = new Point((int)x, (int)y - (int)size.Height - 1);

                        // Define BoundingBox options
                        Pen pen = new Pen(box.BoxColor, 3.2f);
                        SolidBrush colorBrush = new SolidBrush(box.BoxColor);

                        // Draw text on image 
                        thumbnailGraphic.FillRectangle(colorBrush, (int)x, (int)(y - size.Height - 1), (int)size.Width, (int)size.Height);
                        thumbnailGraphic.DrawString(text, drawFont, fontBrush, atPoint);

                        // Draw bounding box on image
                        thumbnailGraphic.DrawRectangle(pen, x, y, width, height);
                    }
                }

                if (!Directory.Exists(outputImageLocation))
                {
                    Directory.CreateDirectory(outputImageLocation);
                }

                image.Save(Path.Combine(outputImageLocation, imageName));
            }
     
        }

        private static void LogDetectedObjects(string imageName, IList<YoloBoundingBox> boundingBoxes)
        {
            Console.WriteLine($".....The objects in the image {imageName} are detected as below....");

            foreach (var box in boundingBoxes)
            {
                Console.WriteLine($"{box.Label} and its Confidence score: {box.Confidence}");
            }

            Console.WriteLine("");
        }
        private static string GetAbsolutePath(string relativePath)
        {
            FileInfo _dataRoot = new FileInfo(typeof(Program).Assembly.Location);
            string assemblyFolderPath = _dataRoot.Directory.FullName;

            string fullPath = Path.Combine(assemblyFolderPath, relativePath);

            return fullPath;
        }
    }
}
