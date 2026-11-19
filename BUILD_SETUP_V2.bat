@echo off
chcp 65001 >nul
cd /d "C:\Users\lenovo\Desktop\ilber"
echo ============================================================================
echo ILBERTech Platform v2.0 - Kurulum Paketi (Setup.exe + Masaustu Ikonu) Derleniyor...
echo ============================================================================

set PYTHON_PATH=C:\Users\lenovo\AppData\Local\Programs\Python\Python311\python.exe
set SITE_PKGS=C:\Users\lenovo\AppData\Local\Programs\Python\Python311\Lib\site-packages
set CTK_PATH=%SITE_PKGS%\customtkinter

"%PYTHON_PATH%" -m PyInstaller --noconfirm --onedir --windowed --name "ILBERTech_Platform" ^
  --icon "C:\Users\lenovo\Desktop\ilber\ilber_logo.ico" ^
  --distpath "C:\Users\lenovo\Desktop\ilber\dist" ^
  --workpath "C:\Users\lenovo\Desktop\ilber\build" ^
  --paths "C:\Users\lenovo\Desktop\ilber" ^
  --add-data "%CTK_PATH%;customtkinter" ^
  --add-data "C:\Users\lenovo\Desktop\ilber\icons;icons" ^
  --add-data "C:\Users\lenovo\Desktop\ilber\ilber_logo.jpg;." ^
  --add-data "C:\Users\lenovo\Desktop\ilber\ilber_logo.ico;." ^
  --collect-data="whisper" ^
  --collect-data="tiktoken_ext" ^
  --collect-data="_sounddevice_data" ^
  --collect-data="certifi" ^
  --collect-data="pdfplumber" ^
  --hidden-import="ilber_core_utils" ^
  --hidden-import="ilber_ui_dialogs" ^
  --hidden-import="ilber_license" ^
  --hidden-import="ilber_field_project" ^
  --hidden-import="torch" ^
  --hidden-import="torchaudio" ^
  --hidden-import="whisper" ^
  --hidden-import="tiktoken" ^
  --hidden-import="tiktoken_ext" ^
  --hidden-import="tiktoken_ext.openai_public" ^
  --hidden-import="transformers" ^
  --hidden-import="sentencepiece" ^
  --hidden-import="llama_cpp" ^
  --hidden-import="docx" ^
  --hidden-import="srt" ^
  --hidden-import="pdfplumber" ^
  --hidden-import="sounddevice" ^
  --hidden-import="_sounddevice_data" ^
  --hidden-import="scipy.io.wavfile" ^
  --hidden-import="translators" ^
  --exclude-module="fitz" ^
  --exclude-module="pymupdf" ^
  --exclude-module="onnxruntime" ^
  --exclude-module="rapidocr_onnxruntime" ^
  --exclude-module="winocr" ^
  --exclude-module="winrt" ^
  --exclude-module="torchvision" ^
  --exclude-module="matplotlib" ^
  --exclude-module="pandas" ^
  --exclude-module="numba" ^
  --exclude-module="llvmlite" ^
  --exclude-module="openpyxl" ^
  --exclude-module="tensorboard" ^
  --exclude-module="notebook" ^
  --exclude-module="IPython" ^
  --exclude-module="pytest" ^
  "C:\Users\lenovo\Desktop\ilber\ilber_v2.py"

echo Kutuphaneler, OCR, PDF/Excel Motoru, Cebir Motoru ve Ikonlar Kopyalaniyor...
xcopy "%SITE_PKGS%\llama_cpp\lib\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\llama_cpp\lib\" /Y /E /I >nul
xcopy "%SITE_PKGS%\whisper\assets\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\whisper\assets\" /Y /E /I >nul
xcopy "%SITE_PKGS%\numba\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\numba\" /Y /E /I >nul
xcopy "%SITE_PKGS%\llvmlite\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\llvmlite\" /Y /E /I >nul
xcopy "%SITE_PKGS%\llvmlite.libs\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\llvmlite.libs\" /Y /E /I >nul 2>&1
xcopy "%SITE_PKGS%\sympy\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\sympy\" /Y /E /I >nul
xcopy "%SITE_PKGS%\mpmath\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\mpmath\" /Y /E /I >nul
xcopy "%SITE_PKGS%\rapidocr_onnxruntime\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\rapidocr_onnxruntime\" /Y /E /I >nul
xcopy "%SITE_PKGS%\onnxruntime\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\onnxruntime\" /Y /E /I >nul
xcopy "%SITE_PKGS%\cv2\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\cv2\" /Y /E /I >nul
xcopy "%SITE_PKGS%\pyclipper\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\pyclipper\" /Y /E /I >nul
xcopy "%SITE_PKGS%\shapely\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\shapely\" /Y /E /I >nul
xcopy "%SITE_PKGS%\shapely.libs\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\shapely.libs\" /Y /E /I >nul 2>&1
xcopy "%SITE_PKGS%\winocr\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\winocr\" /Y /E /I >nul
xcopy "%SITE_PKGS%\winrt\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\winrt\" /Y /E /I >nul
xcopy "%SITE_PKGS%\yaml\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\yaml\" /Y /E /I >nul
xcopy "%SITE_PKGS%\fitz\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\fitz\" /Y /E /I >nul
xcopy "%SITE_PKGS%\pymupdf\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\pymupdf\" /Y /E /I >nul
xcopy "%SITE_PKGS%\openpyxl\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\openpyxl\" /Y /E /I >nul
xcopy "%SITE_PKGS%\et_xmlfile\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\et_xmlfile\" /Y /E /I >nul
copy /Y "C:\Users\lenovo\Desktop\ilber\ilber_core_utils.py" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\" >nul
copy /Y "C:\Users\lenovo\Desktop\ilber\ilber_ui_dialogs.py" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\" >nul
copy /Y "C:\Users\lenovo\Desktop\ilber\ilber_license.py" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\" >nul
copy /Y "C:\Users\lenovo\Desktop\ilber\ilber_field_project.py" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\_internal\" >nul
xcopy "C:\Users\lenovo\Desktop\ilber\icons\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\icons\" /Y /E /I >nul
copy /Y "C:\Users\lenovo\Desktop\ilber\ilber_logo.jpg" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\ilber_logo.jpg" >nul
copy /Y "C:\Users\lenovo\Desktop\ilber\ilber_logo.ico" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\ilber_logo.ico" >nul

echo FFmpeg Ekleniyor...
if exist "C:\Users\lenovo\Desktop\ffmpeg_ext\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe" (
    xcopy "C:\Users\lenovo\Desktop\ffmpeg_ext\ffmpeg-master-latest-win64-gpl-shared\bin\*" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\" /Y /E /I >nul
) else (
    copy /Y "C:\Users\lenovo\AppData\Local\Programs\ILBERTech_Platform\*.dll" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\" >nul 2>&1
    copy /Y "C:\Users\lenovo\AppData\Local\Programs\ILBERTech_Platform\ff*.exe" "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\" >nul 2>&1
)

echo Inno Setup ile ILBERTech_Platform_Setup_v2.exe Olusturuluyor...
"C:\Users\lenovo\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "C:\Users\lenovo\Desktop\ilber\ILBER_Installer.iss"

echo ============================================================================
echo TAMAMLANDI: C:\Users\lenovo\Desktop\ilber\ILBERTech_Platform_Setup_v2.exe
echo ============================================================================
