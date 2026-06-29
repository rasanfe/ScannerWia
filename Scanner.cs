using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WIA;
using System.ComponentModel;

namespace ScannerWia
{
    /// <summary>
    /// Envoltorio interno sobre un dispositivo WIA concreto (un escáner ya seleccionado). Se encarga
    /// de configurar sus propiedades (resolución, tamaño, color…) y de lanzar el escaneo página a página.
    /// </summary>
    internal class Scanner
    {
        // Estas constantes son los IDENTIFICADORES NUMÉRICOS de las propiedades estándar de WIA
        // (los "Property ID" definidos por Windows). WIA no expone nombres amigables: hay que pedir
        // cada propiedad por su número. Por eso van como cadenas, tal y como las espera la API COM.
        const string WIA_SCAN_COLOR_MODE = "6146";
        const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
        const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
        const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
        const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
        const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
        const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
        const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
        const string WIA_SCAN_CONTRAST_PERCENTS = "6155";

        private readonly DeviceInfo _deviceInfo;
        // Ajustes de escaneo por defecto. 2481×3507 px a 300 DPI ≈ un A4 completo
        // (los valores comentados 1250/1700 eran una variante a menor resolución).
        private int resolution = 300;
        private int width_pixel = 2481;//1250;
        private int height_pixel = 3507;//1700;
        private int color_mode = 1; // 1 = color (valores WIA: 1 color, 2 escala grises, 4 blanco/negro)

        internal Scanner(DeviceInfo deviceInfo)
        {
            this._deviceInfo = deviceInfo;
        }

        /// <summary>
        /// Escanea todas las páginas disponibles en el dispositivo en el formato indicado.
        /// Pensado para alimentadores (ADF): repite hasta que el escáner avisa de que no hay más papel.
        /// </summary>
        /// <param name="imageFormat">GUID de formato WIA (uno de los WIA_FORMAT_* de <see cref="ScannerWia"/>).</param>
        /// <returns>Lista de imágenes escaneadas (una por página).</returns>
        internal List<ImageFile> ScanImages(string imageFormat)
        {
            // Connect() abre la conexión COM con el escáner físico y devuelve el objeto Device.
            var device = this._deviceInfo.Connect();

            // CommonDialog de WIA: nos da ShowTransfer, que ejecuta la transferencia de la imagen
            // ya escaneada desde el dispositivo a memoria.
            CommonDialog dlg = new CommonDialog();

            // Items[1]: de nuevo, colección COM 1-based. El item 1 es la "fuente" de escaneo del equipo.
            var item = device.Items[1];

            List<ImageFile> scannedImages = new List<ImageFile>();

            // Bucle hasta agotar páginas: cada vuelta escanea una hoja. Salimos cuando ShowTransfer
            // no devuelve nada o cuando WIA lanza el error de "bandeja vacía" (ver el catch de abajo).
            while (true)
            {
                try
                {
                    AdjustScannerSettings(item, resolution, 0, 0, width_pixel, height_pixel, 0, 0, color_mode);

                    object scanResult = dlg.ShowTransfer(item, imageFormat, true);

                    if (scanResult != null)
                    {
                        var imageFile = (ImageFile)scanResult;
                        scannedImages.Add(imageFile);
                    }
                    else
                    {
                        break;
                    }
                }
                catch (COMException e)
                {
                    // COMException es la excepción típica del interop COM: envuelve el HRESULT que
                    // devuelve el código nativo. Lo traducimos a mensajes claros según su valor.
                    Console.WriteLine(e.ToString());

                    uint errorCode = (uint)e.ErrorCode;

                    // Gestionamos los HRESULT más habituales de WIA por su código hexadecimal.
                    if (errorCode == 0x80210006)
                    {
                        throw new Exception("El escáner está ocupado o no está listo.");
                    }
                    else if (errorCode == 0x80210064)
                    {
                        throw new Exception("El proceso de escaneo ha sido cancelado.");
                    }
                    else if (errorCode == 0x80210003) // WIA_ERROR_PAPER_EMPTY
                    {
                        // No more pages
                        break;
                    }
                    else
                    {
                        throw new Exception("Se produjo un error no capturado, verifica la consola.");
                    }
                }
            }

            return scannedImages;

        }

        /// <summary>
        /// Vuelca todos los ajustes de escaneo (resolución, recorte, brillo, contraste, color) sobre
        /// el item del escáner antes de transferir la imagen. Es, en esencia, configurar el hardware.
        /// </summary>
        /// <param name="scannnerItem">Item de escaneo del dispositivo (su "fuente").</param>
        /// <param name="scanResolutionDPI">Resolución en DPI, p. ej. 300. Se aplica en horizontal y vertical.</param>
        /// <param name="scanStartLeftPixel">Píxel inicial X del área a escanear (recorte por la izquierda).</param>
        /// <param name="scanStartTopPixel">Píxel inicial Y del área a escanear (recorte por arriba).</param>
        /// <param name="scanWidthPixels">Ancho del área de escaneo en píxeles.</param>
        /// <param name="scanHeightPixels">Alto del área de escaneo en píxeles.</param>
        /// <param name="brightnessPercents">Brillo en porcentaje.</param>
        /// <param name="contrastPercents">Contraste en porcentaje.</param>
        /// <param name="colorMode">Modo de color (1 color, 2 grises, 4 blanco/negro).</param>
        private void AdjustScannerSettings(IItem scannnerItem, int scanResolutionDPI, int scanStartLeftPixel, int scanStartTopPixel, int scanWidthPixels, int scanHeightPixels, int brightnessPercents, int contrastPercents, int colorMode)
        {
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, scanStartLeftPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_START_PIXEL, scanStartTopPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidthPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, scanHeightPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, brightnessPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_CONTRAST_PERCENTS, contrastPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_COLOR_MODE, colorMode);
        }

        /// <summary>
        /// Asigna el valor de una propiedad WIA por su identificador, con un plan B si el escáner
        /// rechaza el valor (caso típico de los DPI).
        /// </summary>
        /// <param name="properties">Colección de propiedades del item de escaneo.</param>
        /// <param name="propName">Identificador numérico de la propiedad (uno de los WIA_* de esta clase).</param>
        /// <param name="propValue">Valor a asignar.</param>
        private void SetWIAProperty(IProperties properties, object propName, object propValue)
        {
            // get_Item / set_Value con 'ref' es la firma que genera el interop COM de WIA:
            // estos métodos esperan los argumentos por referencia (herencia del Automation clásico).
            Property prop = properties.get_Item(ref propName);

            try
            {
                prop.set_Value(ref propValue);
            }
            catch
            {
                // El DPI solo admite los valores concretos que soporta el hardware (los lista
                // SubTypeValues). Si el valor pedido no es válido, caemos aquí y elegimos el primero
                // disponible (el más bajo) para que el escaneo no reviente.
                if (propName.ToString() == WIA_HORIZONTAL_SCAN_RESOLUTION_DPI || propName.ToString() == WIA_VERTICAL_SCAN_RESOLUTION_DPI)
                {
                    foreach (object test in prop.SubTypeValues)
                    {
                        prop.set_Value(test);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Representa el escáner por su nombre (propiedad "Name" del dispositivo WIA).
        /// </summary>
        public override string ToString()
        {
            return (string)this._deviceInfo.Properties["Name"].get_Value();
        }




    }
}