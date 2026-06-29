using WIA;
using Tesseract;
using System.Reflection;
using System.Runtime.InteropServices;



namespace ScannerWia
{
    /// <summary>
    /// Fachada pública de la librería de escaneo. Envuelve <b>WIA</b> (Windows Image Acquisition),
    /// la API COM de Windows para hablar con escáneres y cámaras, y opcionalmente pasa la imagen
    /// por <b>Tesseract</b> (OCR) para convertirla en texto.
    /// </summary>
    /// <remarks>
    /// PENSADO PARA CONSUMIRSE DESDE POWERBUILDER: instanciáis <see cref="ScannerWia"/>, llamáis a
    /// <see cref="ListScanners"/> para ver los dispositivos y a <see cref="Scan(string, string, string, string)"/>
    /// para escanear. Los métodos devuelven tipos simples (string[]) precisamente para que el puente
    /// COM/.NET → PowerBuilder sea cómodo.
    ///
    /// ⚠️ POR QUÉ ESTE PROYECTO SOLO COMPILA DESDE VISUAL STUDIO (no por <c>dotnet build</c>):
    /// la referencia a WIA es un <c>&lt;COMReference&gt;</c> en el .csproj. Para usar una librería COM
    /// desde .NET hay que generar un "Interop Assembly" (un envoltorio gestionado sobre el objeto COM)
    /// mediante la herramienta <c>tlbimp</c>. Esa herramienta forma parte de Visual Studio, no del SDK
    /// de .NET, así que el build por CLI falla con el error <c>MSB4803</c> ("COMReference no soportado").
    /// Resumen: WIA es interop COM clásico → necesita Visual Studio para resolverse.
    /// </remarks>
    public class ScannerWia
    {
        // GUIDs de los formatos de imagen que entiende WIA. Son constantes COM bien conocidas
        // (CLSID de cada códec); WIA los espera como cadena para indicarle en qué formato quieres
        // recibir la imagen escaneada.
        const string WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
        const string WIA_FORMAT_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";
        private string _errorText = "";
        private int pageIndex = 1;

        private DeviceManager deviceManager;

        public ScannerWia()
        {
            // DeviceManager es el objeto COM raíz de WIA: a través de él se enumeran los
            // dispositivos de imagen instalados en el equipo.
            deviceManager = new DeviceManager();
        }

        /// <summary>
        /// Localiza un escáner por su nombre dentro de los dispositivos WIA instalados.
        /// </summary>
        /// <param name="scannerName">Nombre exacto tal y como lo devuelve <see cref="ListScanners"/>.</param>
        /// <returns>El <c>DeviceInfo</c> COM correspondiente, o <c>null</c> si no se encuentra.</returns>
        internal DeviceInfo GetScannerByName(string scannerName)
        {
            // Buscar el escáner por nombre dentro de la lista de dispositivos
            return deviceManager.DeviceInfos
                .Cast<DeviceInfo>()
                .FirstOrDefault(deviceInfo => ((DeviceInfo)deviceInfo).Properties["Name"].get_Value().ToString() == scannerName)!;
        }
        /// <summary>
        /// Devuelve los nombres de los escáneres disponibles en el equipo (sin duplicados).
        /// Es lo primero que llamaríais desde PowerBuilder para presentar al usuario la lista.
        /// </summary>
        /// <returns>Array de nombres de escáner; vacío si no hay ninguno.</returns>
        public string[] ListScanners()
        {
            try
            {
                var deviceManager = new DeviceManager();
                var uniqueScannerNames = new HashSet<string>();

                // OJO: las colecciones COM de WIA son 1-based (empiezan en 1, no en 0), por eso el
                // bucle va de 1 hasta Count incluido. Es una herencia del mundo COM/Automation.
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


        /// <summary>
        /// Escanea una o varias páginas desde el escáner indicado y las guarda en disco. Si el
        /// formato es <c>"OCR"</c>, además convierte cada página a un fichero <c>.txt</c> con Tesseract.
        /// </summary>
        /// <param name="scanner">Nombre del escáner (uno de los de <see cref="ListScanners"/>).</param>
        /// <param name="format">Formato de salida: PNG, JPEG, BMP, GIF, TIFF u OCR.</param>
        /// <param name="outputPath">Carpeta de destino donde se guardan los ficheros.</param>
        /// <param name="fileName">Nombre base; se le añade <c>_{nºpágina}</c> y la extensión.</param>
        /// <returns>Rutas de los ficheros generados (imágenes o, en modo OCR, los .txt).</returns>
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
                        // Tesseract necesita una carpeta 'tessdata' con los datos del idioma (.traineddata).
                        // La buscamos junto al ensamblado en ejecución para que la librería funcione
                        // estés donde estés desplegado, sin rutas absolutas codificadas a fuego.
                        string rutaEnsamblado = Assembly.GetExecutingAssembly().Location;
                        string directorio = Path.GetDirectoryName(rutaEnsamblado)!;
                        string dataPath = Path.Combine(directorio, "tessdata");
                        string language = "spa"; // idioma del OCR: español

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

        /// <summary>
        /// Aplica OCR (reconocimiento óptico de caracteres) a una imagen con Tesseract y vuelca el
        /// texto reconocido en un fichero .txt.
        /// </summary>
        /// <param name="imagePath">Imagen de entrada (PNG escaneado).</param>
        /// <param name="txtPath">Fichero de texto de salida.</param>
        /// <param name="dataPath">Carpeta 'tessdata' con los datos de idioma de Tesseract.</param>
        /// <param name="language">Código de idioma, p. ej. "spa" (español) o "eng" (inglés).</param>
        public void ConvertImageToTxt(string imagePath, string txtPath, string dataPath, string language)
        {
            try
            {
                // 'Pix' es el tipo de imagen propio de Tesseract (heredado de la librería Leptonica
                // sobre la que se apoya). Cargamos la imagen, la procesamos y extraemos el texto.
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

        /// <summary>
        /// Devuelve el texto del último error registrado. Útil desde PowerBuilder para mostrar el
        /// motivo de un fallo sin tener que capturar la excepción .NET.
        /// </summary>
        public string GetErrorText()
        {
            return _errorText;
        }

        /// <summary>
        /// Número de páginas escaneadas hasta el momento (el contador interno arranca en 1).
        /// </summary>
        public int GetPageCount()
        {
            return pageIndex - 1;
        }




    }
}
