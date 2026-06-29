# 🖼️ ScannerWia

![.NET](https://img.shields.io/badge/.NET-10.0--windows-512BD4?style=flat-square&logo=dotnet&logoColor=white)
![C#](https://img.shields.io/badge/C%23-239120?style=flat-square&logo=csharp&logoColor=white)
![WIA](https://img.shields.io/badge/Windows-WIA%20(COM)-0078D6?style=flat-square&logo=windows&logoColor=white)
![Blog](https://img.shields.io/badge/blog-rsrsystem-FF5722?style=flat-square&logo=blogger&logoColor=white)

> Librería **.NET 10** para manejar el **escáner** desde PowerBuilder usando **WIA** (Windows Image Acquisition).

## 📋 ¿Qué es esto?

Una librería de clases que habla con el escáner a través de **WIA** (vía interop COM) y, opcionalmente,
pasa lo escaneado por **Tesseract** para OCR. Como el ejemplo original era una app de Windows Forms,
he recreado su interfaz **en PowerBuilder** y dejado aquí solo el motor en C#.

## 🧩 Dependencias

| Paquete / Referencia | Versión |
|----------------------|---------|
| [Tesseract](https://www.nuget.org/packages/Tesseract) | `5.2.0` |
| `Interop.WIA.dll` (interop de WIA **pre-generado con `tlbimp`**, junto al `.csproj`) | — |

## 🔧 Compila FUERA de Visual Studio (truco del interop)

WIA es COM. **Antes** el proyecto usaba `<COMReference>`, que obliga a ejecutar `tlbimp` en cada build →
`dotnet build` fallaba con **`MSB4803`** y solo compilaba desde Visual Studio. **Ya está resuelto:** se
genera **una sola vez** el interop **`Interop.WIA.dll`** con `tlbimp` (desde `C:\Windows\System32\wiaaut.dll`,
namespace `WIA`) y se referencia como `<Reference>` normal con `EmbedInteropTypes=true` + `Private=false`.

Así **compila y publica con `dotnet` (CLI), como el resto** — sin Visual Studio y sin `MSB4803`. Los tipos
WIA quedan **incrustados** en el DLL del proyecto, y `Interop.WIA.dll` **no** hace falta en runtime/despliegue.

```bash
dotnet build ScannerWia.csproj -c Release   # 0 warnings, 0 errores, sin MSB4803
```

> El **código fuente no cambia** (sigue `using WIA;`). Si algún día cambia la versión de WIA del sistema,
> regenera el interop:
> `tlbimp "C:\Windows\System32\wiaaut.dll" /out:Interop.WIA.dll /namespace:WIA`

## 🛠️ Requisitos

- **.NET SDK 10.0** (compila por CLI, **sin Visual Studio**)
- **Windows** con un escáner compatible con **WIA**

## 🔗 Proyecto PowerBuilder relacionado

👉 **pbScanner** — https://github.com/rasanfe/pbScanner

## 🙌 Créditos

Inspirado en el artículo y repositorio de **Our Code World**:

- 📄 https://ourcodeworld.com/articles/read/382/creating-a-scanning-application-in-winforms-with-csharp
- 💻 https://github.com/ourcodeworld/csharp-scanner-wia

---

📨 **Blog:** <https://rsrsystem.blogspot.com/>

> ¡Nos vemos en el próximo artículo! Y recuerda: en PowerBuilder, los límites solo están en nuestra imaginación. 🚀
