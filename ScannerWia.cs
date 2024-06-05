using WIA;
using Tesseract;
using System.Reflection;
using System.Runtime.InteropServices;



namespace ScannerWia
{
    public class ScannerWia
    {
        const string WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";
        private string _errorText="";
        private int pageIndex = 1;

        private DeviceManager deviceManager;

        public ScannerWia()
        {
            deviceManager = new DeviceManager();
        }

        internal DeviceInfo GetScannerByName(string scannerName)
        {
            // Buscar el escáner por nombre dentro de la lista de dispositivos
            return deviceManager.DeviceInfos
                .Cast<DeviceInfo>()
                .FirstOrDefault(deviceInfo => ((DeviceInfo)deviceInfo).Properties["Name"].get_Value().ToString() == scannerName);
        }
        public string[] ListScanners()
        {
            try
            {
                var deviceManager = new DeviceManager();
                var uniqueScannerNames = new HashSet<string>();

                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
                {
                    if (deviceManager.DeviceInfos[i].Type == WiaDeviceType.ScannerDeviceType)
                    {
                        string scannerName = deviceManager.DeviceInfos[i].Properties["Name"].get_Value().ToString();
                        uniqueScannerNames.Add(scannerName);
                    }
                }

                return uniqueScannerNames.ToArray();
            }
            catch (Exception ex)
            {
                _errorText = ex.Message;
                throw new Exception(_errorText);
            }
        }


        public string[] Scan(string scanner, string format, string outputPath, string fileName)
        {

            try
            {
                _errorText = "";
                DeviceInfo deviceInfo = GetScannerByName(scanner);

                var device = new Scanner(deviceInfo);

                if (device == null)
                {
                    _errorText = "Se debe proporcionar un dispositivo de escáner.";
                    throw new ArgumentNullException("device", _errorText);
                }
                else if (string.IsNullOrEmpty(fileName))
                {
                    _errorText = "Se debe proporcionar un nombre de archivo.";
                    throw new ArgumentException(_errorText, "fileName");
                }


                //ImageFile image = new ImageFile();
                List<ImageFile> images = new List<ImageFile>();
                List<string> imagePaths = new List<string>();
                string imageExtension = "";
                string wiaFormat = "";

                switch (format.ToUpper())
                {
                    case "PNG":
                        wiaFormat = WIA_FORMAT_PNG;
                        imageExtension = ".png";
                        break;
                    case "JPEG":
                        wiaFormat = WIA_FORMAT_JPEG;
                        imageExtension = ".jpeg";
                        break;
                    case "BMP":
                        wiaFormat = WIA_FORMAT_BMP;
                        imageExtension = ".bmp";
                        break;
                    case "GIF":
                        wiaFormat = WIA_FORMAT_GIF;
                        imageExtension = ".gif";
                        break;
                    case "TIFF":
                        wiaFormat = WIA_FORMAT_TIFF;
                        imageExtension = ".tiff";
                        break;
                    case "OCR":
                        wiaFormat = WIA_FORMAT_PNG;
                        imageExtension = ".png";
                        break;
                    default:
                        _errorText = "Formato no válido. Formatos admitidos: PNG, JPEG, BMP, GIF, TIFF";
                        throw new ArgumentException(_errorText, "formato");

                }

                // Scan multiple pages
                images = device.ScanImages(wiaFormat);


                // Save single images
                 string singleImagePath = "";

                foreach (var image in images)
                {
                    singleImagePath = Path.Combine(outputPath, $"{fileName}_{pageIndex}{imageExtension}");

                    if (File.Exists(singleImagePath))
                    {
                        File.Delete(singleImagePath);
                    }

                    image.SaveFile(singleImagePath);
                    pageIndex++;

                    //Caso especial para hacer OCR
                    if (format.ToUpper() == "OCR")
                    {
                        //Obtenemos el Nombre del TXT a partir del Nombre del PNG
                        string txtPath = Path.GetDirectoryName(singleImagePath) + "\\" + Path.GetFileNameWithoutExtension(singleImagePath) + ".txt";

                        if (File.Exists(txtPath))
                        {
                            File.Delete(txtPath);
                        }

                        //Convertimos el PNG a TXT
                        string rutaEnsamblado = Assembly.GetExecutingAssembly().Location;
                        string directorio = Path.GetDirectoryName(rutaEnsamblado);
                        string dataPath = Path.Combine(directorio, "tessdata");
                        string language = "spa";

                        ConvertImageToTxt(singleImagePath, txtPath, dataPath, language);

                        //Eliminamos Imagen Escaneada (PNG)
                        File.Delete(singleImagePath);

                        //Cambio la Ruta de la Imagen por la del TXT
                        singleImagePath = txtPath;

                    }
                    imagePaths.Add(singleImagePath);


                }
                return imagePaths.ToArray();
            }
            catch (Exception ex) 
            {
                _errorText = ex.Message;
                throw new Exception(_errorText);
            }
        }

        public void ConvertImageToTxt(string imagePath, string txtPath, string dataPath, string language)
        {
            try
            {
                var engine = new TesseractEngine(@dataPath, language);
                var image = Pix.LoadFromFile(@imagePath);
                var page = engine.Process(image);

                var text = page.GetText();

                File.WriteAllText(@txtPath, text);
            }
            catch (Exception ex)
            {
                _errorText = ex.Message;
                throw new Exception(_errorText);
            }
        }

        public string GetErrorText()
        {
            return _errorText;
        }

        public int GetPageCount()
        {
            return pageIndex - 1;
        }




    }
}
