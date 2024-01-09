using System;
using System.Collections.Generic;
using System.Linq;
using WIA;
using System.IO;
using Tesseract;
using System.Reflection;

namespace ScannerWia
{
    public class ScannerWia
    {
        public string errorText = "";
        const string WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";

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


        public string Scan(string scanner, string format, string outputPath, string fileName)
        {

            DeviceInfo deviceInfo = GetScannerByName(scanner);

            var device = new Scanner(deviceInfo);

            if (device == null)
            {
                throw new ArgumentNullException("device", "Se debe proporcionar un dispositivo de escáner.");
            }
            else if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("Se debe proporcionar un nombre de archivo.", "fileName");
            }


            ImageFile image = new ImageFile();
            string imageExtension = "";

            switch (format.ToUpper())
            {
                case "PNG":
                    image = device.ScanImage(WIA_FORMAT_PNG);
                    imageExtension = ".png";
                    break;
                case "JPEG":
                    image = device.ScanImage(WIA_FORMAT_JPEG);
                    imageExtension = ".jpeg";
                    break;
                case "BMP":
                    image = device.ScanImage(WIA_FORMAT_BMP);
                    imageExtension = ".bmp";
                    break;
                case "GIF":
                    image = device.ScanImage(WIA_FORMAT_GIF);
                    imageExtension = ".gif";
                    break;
                case "TIFF":
                    image = device.ScanImage(WIA_FORMAT_TIFF);
                    imageExtension = ".tiff";
                    break;
                case "OCR":
                    image = device.ScanImage(WIA_FORMAT_PNG);
                    imageExtension = ".png";
                    break;
                default:
                    throw new ArgumentException("Formato no válido. Formatos admitidos: PNG, JPEG, BMP, GIF, TIFF", "formato");

            }

            // Save the image
            var path = Path.Combine(outputPath, fileName + imageExtension);

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            image.SaveFile(path);

            //Caso especial para hacer OCR
            if (format.ToUpper() == "OCR")
            {
                //Obtenemos el Nombre del TXT a partir del Nombre del PNG
                string txtPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + ".txt";

                if (File.Exists(txtPath))
                {
                    File.Delete(txtPath);
                }

                //Convertimos el PNG a TXT
                string rutaEnsamblado = Assembly.GetExecutingAssembly().Location;
                string directorio = Path.GetDirectoryName(rutaEnsamblado);
                string dataPath = Path.Combine(directorio, "tessdata");
                string language = "spa";

                ConvertImageToTxt(path, txtPath, dataPath, language);

                //Eliminamos Imagen Escaneada (PNG)
                File.Delete(path);

                //Cambio la Ruta de la Imagen por la del TXT
                path = txtPath;

            }

            return path;
        }

        public void ConvertImageToTxt(string imagePath, string txtPath, string dataPath, string language)
        {
            var engine = new TesseractEngine(@dataPath, language);
            var image = Pix.LoadFromFile(@imagePath);
            var page = engine.Process(image);

            var text = page.GetText();

            File.WriteAllText(@txtPath, text);
        }



    }
}
