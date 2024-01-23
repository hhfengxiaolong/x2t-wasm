/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

#include "ASCConverters.h"
//todo убрать ошибки компиляции если переместить include ниже
// #include "../../PdfWriter/OnlineOfficeBinToPdf.h"
#include "cextracttools.h"

#include "../../OfficeUtils/src/OfficeUtils.h"
#include "../../Common/3dParty/pole/pole.h"

// #include "../../ASCOfficeDocxFile2/DocWrapper/DocxSerializer.h"
// #include "../../ASCOfficeDocxFile2/DocWrapper/XlsxSerializer.h"
#include "../../ASCOfficePPTXFile/ASCOfficePPTXFile.h"
// #include "../../ASCOfficeRtfFile/RtfFormatLib/source/ConvertationManager.h"
#include "../../ASCOfficeDocFile/DocFormatLib/DocFormatLib.h"
// #include "../../ASCOfficeTxtFile/TxtXmlFormatLib/Source/TxtXmlFile.h"
#include "../../ASCOfficePPTFile/PPTFormatLib/PPTFormatLib.h"
// #include "../../ASCOfficeOdfFile/src/ConvertOO2OOX.h"
// #include "../../ASCOfficeOdfFileW/source/Oox2OdfConverter/Oox2OdfConverter.h"
// #include "../../DesktopEditor/doctrenderer/doctrenderer.h"
// #include "../../DesktopEditor/doctrenderer/docbuilder.h"
#include "../../DesktopEditor/graphics/pro/Fonts.h"
// #include "../../DesktopEditor/graphics/MetafileToGraphicsRenderer.h"
// #include "../../PdfReader/PdfReader.h"
#include "../../PdfReader/Src/ErrorConstants.h"
// #include "../../DjVuFile/DjVu.h"
// #include "../../XpsFile/XpsFile.h"
// #include "../../HtmlRenderer/include/HTMLRenderer3.h"
// #include "../../HtmlFile/HtmlFile.h"
// #include "../../ASCOfficeXlsFile2/source/XlsXlsxConverter/ConvertXls2Xlsx.h"
// #include "../../OfficeCryptReader/source/ECMACryptFile.h"

#include "../../DesktopEditor/common/Path.h"
#include "../../DesktopEditor/common/Directory.h"

#include <iostream>
#include <fstream>

namespace NExtractTools
{
    void initApplicationFonts(NSFonts::IApplicationFonts* pApplicationFonts, InputParams& params)
    {
		std::wstring sFontPath = params.getFontPath();
        
		if(sFontPath.empty())
            pApplicationFonts->Initialize();
        else
            pApplicationFonts->InitializeFromFolder(sFontPath);
    }
	std::wstring getExtentionByRasterFormat(int format)
	{
		std::wstring sExt;
		switch(format)
		{
			case 1:
				sExt = L".bmp";
			break;
			case 2:
				sExt = L".gif";
			break;
			case 3:
				sExt = L".jpg";
			break;
			default:
				sExt = L".png";
			break;
		}
		return sExt;
	}
	_UINT32 replaceContentType(const std::wstring &sDir, const std::wstring &sCTFrom, const std::wstring &sCTTo)
	{
        _UINT32 nRes = 0;
		std::wstring sContentTypesPath = sDir + FILE_SEPARATOR_STR + _T("[Content_Types].xml");
		if (NSFile::CFileBinary::Exists(sContentTypesPath))
		{
			std::wstring sData;
			if (NSFile::CFileBinary::ReadAllTextUtf8(sContentTypesPath, sData))
			{
				sData = string_replaceAll(sData, sCTFrom, sCTTo);
				if (false == NSFile::CFileBinary::SaveToFile(sContentTypesPath, sData, true))
					nRes = AVS_FILEUTILS_ERROR_CONVERT;
			}
		}
		return nRes;
	}
	_UINT32 processEncryptionError(_UINT32 hRes, const std::wstring &sFrom, InputParams& params)
	{
		if (AVS_ERROR_DRM == hRes)
		{
			if(!params.getDontSaveAdditional())
			{
				copyOrigin(sFrom, *params.m_sFileTo);
			}
			return AVS_FILEUTILS_ERROR_CONVERT_DRM;
		}
		else if (AVS_ERROR_PASSWORD == hRes)
		{
			return AVS_FILEUTILS_ERROR_CONVERT_PASSWORD;
		}
		return hRes;
	}
	// docx -> bin
    _UINT32 docx2doct_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return 0;
    }
    _UINT32 docx_dir2doct_bin (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
    {
        return 0;
    }

    // docx -> doct
    _UINT32 docx2doct (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }

	// bin -> docx
    _UINT32 doct_bin2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
    // bin -> docx dir
    _UINT32 doct_bin2docx_dir (const std::wstring &sFrom, const std::wstring &sToResult, const std::wstring &sTo, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }

    // doct -> docx
    _UINT32 doct2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
	// dotx -> docx
	_UINT32 dotx2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 dotx2docx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	// docm -> docx
	_UINT32 docm2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 docm2docx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return 0;
	}
	// dotm -> docx
	_UINT32 dotm2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 dotm2docx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return 0;
	}
	// dotm -> docm
	_UINT32 dotm2docm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 dotm2docm_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
    // xslx -> bin
    _UINT32 xlsx2xlst_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
		return AVS_FILEUTILS_ERROR_CONVERT;
    }
	_UINT32 xlsx_dir2xlst_bin (const std::wstring &sXlsxDir, const std::wstring &sTo, InputParams& params, bool bXmlOptions, const std::wstring &sXlsxFile)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }

    // xslx -> xslt
    _UINT32 xlsx2xlst (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }

    // bin -> xslx
    _UINT32 xlst_bin2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
    _UINT32 xlst_bin2xlsx_dir (const std::wstring &sFrom, const std::wstring &sToResult, const std::wstring &sTo, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }

    // xslt -> xslx
    _UINT32 xlst2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
	// xltx -> xlsx
	_UINT32 xltx2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 xltx2xlsx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	// xlsm -> xlsx
	_UINT32 xlsm2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 xlsm2xlsx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
		return 0;
	}
	// xltm -> xlsx
	_UINT32 xltm2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 xltm2xlsx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
		return 0;
	}
	// xltm -> xlsm
	_UINT32 xltm2xlsm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 xltm2xlsm_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
    // pptx -> bin
    _UINT32 pptx2pptt_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return 0;
	}
    _UINT32 pptx_dir2pptt_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return 0;
	}
    // pptx -> pptt
    _UINT32 pptx2pptt (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return 0;
    }

    // bin -> pptx
    _UINT32 pptt_bin2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return 0;
	}
    _UINT32 pptt_bin2pptx_dir (const std::wstring &sFrom, const std::wstring &sToResult, const std::wstring &sTo, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return 0;
	}
	// pptt -> pptx
    _UINT32 pptt2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
    {
        return 0;
    }
    // zip dir
    _UINT32 dir2zip (const std::wstring &sFrom, const std::wstring &sTo)
    {
        COfficeUtils oCOfficeUtils(NULL);
        return (S_OK == oCOfficeUtils.CompressFileOrDirectory(sFrom, sTo)) ? 0 : AVS_FILEUTILS_ERROR_CONVERT;
    }

    // unzip dir
    _UINT32 zip2dir (const std::wstring &sFrom, const std::wstring &sTo)
    {
        COfficeUtils oCOfficeUtils(NULL);
        return (S_OK == oCOfficeUtils.ExtractToDirectory(sFrom, sTo, NULL, 0)) ? 0 : AVS_FILEUTILS_ERROR_CONVERT;
    }

    // csv -> xslt
    _UINT32 csv2xlst (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }

    // csv -> xslx
    _UINT32 csv2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
    _UINT32 csv2xlst_bin (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
    // xlst -> csv
	_UINT32 xlst2csv (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
    {
        return AVS_FILEUTILS_ERROR_CONVERT;
    }
	// xslx -> csv
	_UINT32 xlsx2csv (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 xlst_bin2csv (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
   }
	// bin -> pdf
	_UINT32 bin2pdf (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bPaid, const std::wstring &sThemeDir, InputParams& params)
	{
        return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 bin2image (const std::wstring &sTFileDir, BYTE* pBuffer, LONG lBufferLen, const std::wstring &sTo, const std::wstring &sTemp, const std::wstring &sThemeDir, InputParams& params)
	{
		return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 bin2imageBase64 (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, const std::wstring &sThemeDir, InputParams& params)
	{
		return AVS_FILEUTILS_ERROR_CONVERT;
	}
   //doct_bin -> pdf
	_UINT32 doct_bin2pdf(NSDoctRenderer::DoctRendererFormat::FormatFile eFromType, const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bPaid, const std::wstring &sThemeDir, InputParams& params)
   {
       return AVS_FILEUTILS_ERROR_CONVERT;
   }

	//doct_bin -> image
	_UINT32 doct_bin2image(NSDoctRenderer::DoctRendererFormat::FormatFile eFromType, const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bPaid, const std::wstring &sThemeDir, InputParams& params)
	{
		return AVS_FILEUTILS_ERROR_CONVERT;
	}

	// ppsx -> pptx
	_UINT32 ppsx2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 ppsx2pptx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
   	}
	// pptm -> pptx
	_UINT32 pptm2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 pptm2pptx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return 0;
   }
	// potm -> pptx
	_UINT32 potm2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 potm2pptx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return 0;
   }
	// ppsm -> pptx
	_UINT32 ppsm2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 ppsm2pptx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return 0;
   }
	// potx -> pptx
	_UINT32 potx2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 potx2pptx_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	// potm -> pptm
	_UINT32 potm2pptm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 potm2pptm_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	// ppsm -> pptm
	_UINT32 ppsm2pptm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 ppsm2pptm_dir (const std::wstring &sFrom, const std::wstring &sTo, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	// ppt -> pptx
	_UINT32 ppt2pptx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       std::wstring sResultPptxDir = sTemp + FILE_SEPARATOR_STR + _T("pptx_unpacked");

       NSDirectory::CreateDirectory(sResultPptxDir);

       _UINT32 nRes = ppt2pptx_dir(sFrom, sResultPptxDir, sTemp, params);

		nRes = processEncryptionError(nRes, sFrom, params);
		if(SUCCEEDED_X2T(nRes))
		{
           COfficeUtils oCOfficeUtils(NULL);
           if(S_OK == oCOfficeUtils.CompressFileOrDirectory(sResultPptxDir, sTo, true))
               return 0;
		}	
		return nRes;
	}
	_UINT32 ppt2pptx_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		COfficePPTFile pptFile;

		pptFile.put_TempDirectory(sTemp);
	   
		bool bMacros = false;
		long nRes = pptFile.LoadFromFile(sFrom, sTo, params.getPassword(), bMacros);
		nRes = processEncryptionError(nRes, sFrom, params);
		return nRes;
	}
	// ppt -> pptm
	_UINT32 ppt2pptm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 ppt2pptm_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
	}
	// ppt -> pptt
	_UINT32 ppt2pptt (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return 0;
   }
	// ppt -> pptt_bin
	_UINT32 ppt2pptt_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
        return 0;
   }

	// pptx -> odp
	_UINT32 pptx2odp (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params )
	{
        return AVS_FILEUTILS_ERROR_CONVERT;
	}
	// pptx_dir -> odp
	_UINT32 pptx_dir2odp (const std::wstring &sPptxDir, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params, bool bTemplate)
	{
		return 0;
	}
	// rtf -> docx
	_UINT32 rtf2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
   }
	_UINT32 rtf2docx_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
        return 0;
   }

	// rtf -> doct
	_UINT32 rtf2doct (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return 0;
   }

	// rtf -> doct_bin
	_UINT32 rtf2doct_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
        return AVS_FILEUTILS_ERROR_CONVERT;
   }

	// docx -> rtf
	_UINT32 docx2rtf (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
		return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 docx_dir2rtf(const std::wstring &sDocxDir, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
   }

	// doc -> docx
	_UINT32 doc2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       std::wstring sResultDocxDir = sTemp + FILE_SEPARATOR_STR + _T("docx_unpacked");

       NSDirectory::CreateDirectory(sResultDocxDir);

       _UINT32 hRes = doc2docx_dir(sFrom, sResultDocxDir, sTemp, params);
       if(SUCCEEDED_X2T(hRes))
       {
           COfficeUtils oCOfficeUtils(NULL);
           if(S_OK == oCOfficeUtils.CompressFileOrDirectory(sResultDocxDir, sTo, true))
               return 0;
       }
       else if (AVS_ERROR_DRM == hRes)
       {
           if(!params.getDontSaveAdditional())
           {
               copyOrigin(sFrom, *params.m_sFileTo);
           }
           return AVS_FILEUTILS_ERROR_CONVERT_DRM;
       }
       else if (AVS_ERROR_PASSWORD == hRes)
       {
           return AVS_FILEUTILS_ERROR_CONVERT_PASSWORD;
       }
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 doc2docx_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
        COfficeDocFile docFile;
		docFile.m_sTempFolder = sTemp;
		
		bool bMacros = false;

		_UINT32 hRes = docFile.LoadFromFile( sFrom, sTo, params.getPassword(), bMacros, NULL);
		if (AVS_ERROR_DRM == hRes)
		{
			if(!params.getDontSaveAdditional())
			{
				copyOrigin(sFrom, *params.m_sFileTo);
			}
			return AVS_FILEUTILS_ERROR_CONVERT_DRM;
		}
		else if (AVS_ERROR_PASSWORD == hRes)
		{
			return AVS_FILEUTILS_ERROR_CONVERT_PASSWORD;
		}
		return 0 == hRes ? 0 : AVS_FILEUTILS_ERROR_CONVERT;
   }

	// doc -> docm
	_UINT32 doc2docm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 doc2docm_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
   }

	// doc -> doct
	_UINT32 doc2doct (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return 0;
	}

	// doc -> doct_bin
	_UINT32 doc2doct_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
        return 0;
	}
	_UINT32 docx_dir2doc (const std::wstring &sDocxDir, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return /*S_OK == docFile.SaveToFile(sTo, sDocxDir, NULL) ? 0 : */AVS_FILEUTILS_ERROR_CONVERT;
	}

	// doct -> rtf
	_UINT32 doct2rtf (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
	{
       return AVS_FILEUTILS_ERROR_CONVERT;
   }

	// bin -> rtf
	_UINT32 doct_bin2rtf (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bFromChanges, const std::wstring &sThemeDir, InputParams& params)
   {
       return 0;
   }
	// txt -> docx
	_UINT32 txt2docx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
       return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 txt2docx_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
	   return 0;
   }
	// txt -> doct
	_UINT32 txt2doct (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return 0;
   }

	// txt -> doct_bin
	_UINT32 txt2doct_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
        return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 docx_dir2txt (const std::wstring &sDocxDir, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
	}
	//odf
	_UINT32 odf2oot(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
       return 0;
   }

	_UINT32 odf2oot_bin(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
   {
		return 0;
	}
	_UINT32 otf2odf(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}

	_UINT32 odf2oox(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 odf2oox_dir(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
	   return 0;
	}
	//odf flat
	_UINT32 odf_flat2oot(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
       return 0;
	}

	_UINT32 odf_flat2oot_bin(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
       return 0;
	}
	_UINT32 odf_flat2oox(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
       return 0;
	}
	_UINT32 odf_flat2oox_dir(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	// docx -> odt
	_UINT32 docx2odt (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params )
   {
        return AVS_FILEUTILS_ERROR_CONVERT;
   }
	// docx dir -> odt
	_UINT32 docx_dir2odt (const std::wstring &sDocxDir, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params, bool bTemplate)
   {
       return 0;
   }
	// xlsx -> ods
	_UINT32 xlsx2ods (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params )
   {
        return AVS_FILEUTILS_ERROR_CONVERT;
   }

	_UINT32 xlsx_dir2ods (const std::wstring &sXlsxDir, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params, bool bTemplate)
   {
       return 0;
	}

	_UINT32 mscrypt2oot (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 mscrypt2oox	 (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 mscrypt2oot_bin	 (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return AVS_FILEUTILS_ERROR_CONVERT;
	}
	_UINT32 oox2mscrypt	 (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
    _UINT32 fromMscrypt (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
        return 0;
	}
 	//html
	_UINT32 html2doct_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
       return 0;
   }
 	//html in container
	_UINT32 html_zip2doct_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
	}
	//mht
	_UINT32 mht2doct_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return 0;
	}
	_UINT32 epub2doct_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
       return 0;
   }
	// mailmerge
	_UINT32 convertmailmerge (const InputParamsMailMerge& oMailMergeSend,const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, bool bPaid, const std::wstring &sThemeDir, InputParams& params)
   {
       return 0;
   }

    // _UINT32 PdfDjvuXpsToImage(IOfficeDrawingFile** ppReader, const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params, NSFonts::IApplicationFonts* pApplicationFonts)
	// {
	// 	return 0;
	// }

	_UINT32 fromDocxDir(const std::wstring &sFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
   {
       return 0;
   }
	_UINT32 fromDoctBin(const std::wstring &sFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
   {
       return 0;
   }
	_UINT32 fromDocument(const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTemp, InputParams& params)
   {
       return 0;
   }

	_UINT32 fromXlsxDir(const std::wstring &sFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params, const std::wstring &sXlsxFile)
   {
       return 0;
   }
	_UINT32 fromSpreadsheet(const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTemp, InputParams& params)
   {
       return 0;
   }

	_UINT32 fromPptxDir(const std::wstring &sFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
	{
		return 0;
	}
	_UINT32 fromPpttBin(const std::wstring &sFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
	{
       return 0;
   }
	_UINT32 fromPresentation(const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTemp, InputParams& params)
   {
       return 0;
   }

	_UINT32 fromT(const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
   {
       return 0;
   }
	_UINT32 fromCrossPlatform(const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
   {
       return 0;
   }
	_UINT32 fromCanvasPdf(const std::wstring &sFrom, int nFormatFrom, const std::wstring &sTo, int nFormatTo, const std::wstring &sTemp, const std::wstring &sThemeDir, bool bFromChanges, bool bPaid, InputParams& params)
   {
       return 0;
   }

	// xls -> xlsx
	_UINT32 xls2xlsx (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
       return AVS_FILEUTILS_ERROR_CONVERT;
   }
	_UINT32 xls2xlsx_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
       return 0;
   }
	// xls -> xlsm
	_UINT32 xls2xlsm (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 xls2xlsm_dir (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
	{
		return 0;
   }
	// xls -> xlst
	_UINT32 xls2xlst (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
       return 0;
   }

	// xls -> xlst_bin
	_UINT32 xls2xlst_bin (const std::wstring &sFrom, const std::wstring &sTo, const std::wstring &sTemp, InputParams& params)
   {
        return 0;
   }
	_UINT32 html2doct_bin(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 html_zip2doct_bin(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 html2doct(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
	_UINT32 html_zip2doct(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}
    _UINT32 html2docx(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}

    _UINT32 html_zip2docx(const std::wstring &sFrom, const std::wstring &sTo, const std::wstring & sTemp, InputParams& params)
	{
		return 0;
	}

//------------------------------------------------------------------------------------------------------------------
	_UINT32 detectMacroInFile(InputParams& oInputParams)
	{
		return 0;
	}
	_UINT32 fromInputParams(InputParams& oInputParams)
	{
		TConversionDirection conversion  = oInputParams.getConversionDirection();
		std::wstring sFileFrom	= *oInputParams.m_sFileFrom;
		std::wstring sFileTo	= *oInputParams.m_sFileTo;

		int nFormatFrom = AVS_OFFICESTUDIO_FILE_UNKNOWN;
		if(NULL != oInputParams.m_nFormatFrom)
			nFormatFrom = *oInputParams.m_nFormatFrom;
		int nFormatTo = AVS_OFFICESTUDIO_FILE_UNKNOWN;
		if(NULL != oInputParams.m_nFormatTo)
			nFormatTo = *oInputParams.m_nFormatTo;

		if (TCD_ERROR == conversion)
		{
			if(AVS_OFFICESTUDIO_FILE_DOCUMENT_TXT == nFormatFrom || AVS_OFFICESTUDIO_FILE_SPREADSHEET_CSV == nFormatFrom)
				return AVS_FILEUTILS_ERROR_CONVERT_NEED_PARAMS;
			else{
				// print out conversion direction error
				std::cerr << "Couldn't recognize conversion direction from an argument" << std::endl;
				return AVS_FILEUTILS_ERROR_CONVERT_PARAMS;
			}
		}

		if (sFileFrom.empty() || sFileTo.empty())
		{
			std::cerr << "Empty sFileFrom or sFileTo" << std::endl;
			return AVS_FILEUTILS_ERROR_CONVERT_PARAMS;
		}

		// get conversion direction from file extensions
		if (TCD_AUTO == conversion)
		{
			conversion = getConversionDirectionFromExt (sFileFrom, sFileTo);
		}

		if (TCD_ERROR == conversion)
		{
			// print out conversion direction error
			std::cerr << "Couldn't automatically recognize conversion direction from extensions" << std::endl;
			return AVS_FILEUTILS_ERROR_CONVERT_PARAMS;
		}
		bool bFromChanges = false;
		if(NULL != oInputParams.m_bFromChanges)
			bFromChanges = *oInputParams.m_bFromChanges;
		bool bPaid = true;
		if(NULL != oInputParams.m_bPaid)
			bPaid = *oInputParams.m_bPaid;
		std::wstring sThemeDir;
		if(NULL != oInputParams.m_sThemeDir)
			sThemeDir = *oInputParams.m_sThemeDir;
		InputParamsMailMerge* oMailMerge = NULL;
		if(NULL != oInputParams.m_oMailMergeSend)
			oMailMerge = oInputParams.m_oMailMergeSend;

		bool bExternalTempDir = false;
		std::wstring sTempDir;
		if (NULL != oInputParams.m_sTempDir)
		{
			bExternalTempDir = true;
			sTempDir = *oInputParams.m_sTempDir;
		}
		else
		{
			sTempDir = NSDirectory::CreateDirectoryWithUniqueName(NSDirectory::GetFolderPath(sFileTo));
		}
		if (sTempDir.empty())
		{
			std::cerr << "Couldn't create temp folder" << std::endl;
			return AVS_FILEUTILS_ERROR_UNKNOWN;
		}

		if (!oInputParams.checkInputLimits())
		{
			return AVS_FILEUTILS_ERROR_CONVERT_LIMITS;
		}

		_UINT32 result = 0;
		switch(conversion)
		{
			case TCD_DOCX2DOCT:
			{
				result = docx2doct (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOCT2DOCX:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_DOCUMENT_DOCX);
				result =  doct2docx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_DOCT2DOTX:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_DOCUMENT_DOTX);
				result =  doct2docx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_DOCT2DOCM:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_DOCUMENT_DOCM);
				result =  doct2docx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_XLSX2XLST:
			{
				result =  xlsx2xlst (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLST2XLSX:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_SPREADSHEET_XLSX);
				result =  xlst2xlsx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_XLST2XLSM:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_SPREADSHEET_XLSM);
				result =  xlst2xlsx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_XLST2XLTX:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_SPREADSHEET_XLTX);
				result =  xlst2xlsx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_PPTX2PPTT:
			{
				result =  pptx2pptt (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPTT2PPTX:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_PRESENTATION_PPTX);
				result =  pptt2pptx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;	
			case TCD_PPTT2PPTM:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_PRESENTATION_PPTM);
				result =  pptt2pptx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;	
			case TCD_PPTT2POTX:
			{
				oInputParams.m_nFormatTo = new int(AVS_OFFICESTUDIO_FILE_PRESENTATION_POTX);
				result =  pptt2pptx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;	
			case TCD_DOTX2DOCX:
			{
				result =  dotx2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOCM2DOCX:
			{
				result =  docm2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOTM2DOCX:
			{
				result =  dotm2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOTM2DOCM:
			{
				result =  dotm2docm (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLTX2XLSX:
			{
				result =  xltx2xlsx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLSM2XLSX:
			{
				result =  xltx2xlsx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLTM2XLSX:
			{
				result =  xltm2xlsx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLTM2XLSM:
			{
				result =  xltm2xlsm (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPSX2PPTX:
			{
				result =  ppsx2pptx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_POTX2PPTX:
			{
				result =  potx2pptx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_POTM2PPTX:
			{
				result =  potm2pptx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPSM2PPTX:
			{
				result =  ppsm2pptx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_POTM2PPTM:
			{
				result =  potm2pptm (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPSM2PPTM:
			{
				result =  ppsm2pptm (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPTM2PPTX:
			{
				result =  pptm2pptx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_ZIPDIR:
			{
				result =  dir2zip (sFileFrom, sFileTo);
			}break;
			case TCD_UNZIPDIR:
			{
				result =  zip2dir (sFileFrom, sFileTo);
			}break;
			case TCD_CSV2XLSX:
			{
				result =  csv2xlsx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_CSV2XLST:
			{
				result =  csv2xlst (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLSX2CSV:
			{
				result =  xlsx2csv (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLST2CSV:
			{
				result =  xlst2csv (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOCX2DOCT_BIN:
			{
				result = docx2doct_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOCT_BIN2DOCX:
			{
				result =  doct_bin2docx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_XLSX2XLST_BIN:
			{
				result =  xlsx2xlst_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLST_BIN2XLSX:
			{
				result =  xlst_bin2xlsx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_PPTX2PPTT_BIN:
			{
				result =  pptx2pptt_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPTT_BIN2PPTX:
			{
				result =  pptt_bin2pptx (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_BIN2PDF:
			{
				result =  bin2pdf (sFileFrom, sFileTo, sTempDir, bPaid, sThemeDir, oInputParams);
			}break;
			case TCD_BIN2T:
			{
                result =  dir2zip (NSDirectory::GetFolderPath(sFileFrom), sFileTo);
			}break;
			case TCD_T2BIN:
			{
                result =  zip2dir (sFileFrom, NSDirectory::GetFolderPath(sFileTo));
			}break;
			case TCD_PPT2PPTX:
			{
				result =  ppt2pptx (sFileFrom, sFileTo, sTempDir, oInputParams);
            }break;
			case TCD_PPT2PPTM:
			{
				result =  ppt2pptm (sFileFrom, sFileTo, sTempDir, oInputParams);
            }break;
			case TCD_PPT2PPTT:
			{
				result =  ppt2pptt (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPT2PPTT_BIN:
			{
				result =  ppt2pptt_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_RTF2DOCX:
			{
				result =  rtf2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_RTF2DOCT:
			{
				result = rtf2doct (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_RTF2DOCT_BIN:
			{
				result = rtf2doct_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOCX2RTF:
			{
				result =  docx2rtf (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOC2DOCX:
			{
				result =  doc2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOC2DOCM:
			{
				result =  doc2docm (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_DOCT2RTF:
			{
				result =  doct2rtf (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_DOCT_BIN2RTF:
			{
				result =  doct_bin2rtf (sFileFrom, sFileTo, sTempDir, bFromChanges, sThemeDir, oInputParams);
			}break;
			case TCD_TXT2DOCX:
			{
				result =  txt2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_TXT2DOCT:
			{
				result = txt2doct (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_TXT2DOCT_BIN:
			{
				result = txt2doct_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLS2XLSX:
			{
				result =  xls2xlsx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLS2XLSM:
			{
				result =  xls2xlsm (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLS2XLST:
			{
				result = xls2xlst (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLS2XLST_BIN:
			{
				result = xls2xlst_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_OTF2ODF:
			{
				result =  otf2odf (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_ODF2OOX:
			{
				result =  odf2oox (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_ODF2OOT:
			{
				result = odf2oot (sFileFrom, sFileTo,  sTempDir, oInputParams);
			}break;
			case TCD_ODF2OOT_BIN:
			{
				result = odf2oot_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
            }break;
			case TCD_ODF_FLAT2OOX:
			{
				result =  odf_flat2oox (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_ODF_FLAT2OOT:
			{
				result = odf_flat2oot (sFileFrom, sFileTo,  sTempDir, oInputParams);
			}break;
			case TCD_ODF_FLAT2OOT_BIN:
			{
				result = odf_flat2oot_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
            }break;
			case TCD_DOCX2ODT:
			{
				result =  docx2odt (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_XLSX2ODS:
			{
				result =  xlsx2ods (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_PPTX2ODP:
			{
				result =  pptx2odp (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_MAILMERGE:
			{
				result = convertmailmerge(*oMailMerge, sFileFrom, sFileTo, sTempDir, bPaid, sThemeDir, oInputParams);
			}break;
			case TCD_DOCUMENT2:
			{
				result = fromDocument(sFileFrom, nFormatFrom, sTempDir, oInputParams);
			}break;
			case TCD_SPREADSHEET2:
			{
				result = fromSpreadsheet(sFileFrom, nFormatFrom, sTempDir, oInputParams);
			}break;
			case TCD_PRESENTATION2:
			{
				result = fromPresentation(sFileFrom, nFormatFrom, sTempDir, oInputParams);
			}break;
			case TCD_T2:
			{
				result = fromT(sFileFrom, nFormatFrom, sFileTo, nFormatTo, sTempDir, sThemeDir, bFromChanges, bPaid, oInputParams);
			}break;
			case TCD_DOCT_BIN2:
			{
				result = fromDoctBin(sFileFrom, sFileTo, nFormatTo, sTempDir, sThemeDir, bFromChanges, bPaid, oInputParams);
			}break;
			case TCD_XLST_BIN2:
			{
				result = fromXlstBin(sFileFrom, sFileTo, nFormatTo, sTempDir, sThemeDir, bFromChanges, bPaid, oInputParams);
			}break;
			case TCD_PPTT_BIN2:
			{
				result = fromPpttBin(sFileFrom, sFileTo, nFormatTo, sTempDir, sThemeDir, bFromChanges, bPaid, oInputParams);
			}break;
			case TCD_CROSSPLATFORM2:
			{
				result = fromCrossPlatform(sFileFrom, nFormatFrom, sFileTo, nFormatTo, sTempDir, sThemeDir, bFromChanges, bPaid, oInputParams);
			}break;
			case TCD_CANVAS_PDF2:
			{
				result = fromCanvasPdf(sFileFrom, nFormatFrom, sFileTo, nFormatTo, sTempDir, sThemeDir, bFromChanges, bPaid, oInputParams);
			}break;
			case TCD_MSCRYPT2:
			{
				result = fromMscrypt (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_MSCRYPT2_RAW:
			{
				result = mscrypt2oox(sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_2MSCRYPT_RAW:
			{
				result = oox2mscrypt(sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_MSCRYPT2DOCT:
			case TCD_MSCRYPT2XLST:
			case TCD_MSCRYPT2PPTT:
			{
				result = mscrypt2oot (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_MSCRYPT2BIN:
				result =  mscrypt2oot_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			{
			}break;
			case TCD_HTML2DOCX:
			{
				result = html2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_HTMLZIP2DOCX:
			{
				result = html_zip2docx (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_HTML2DOCT:
			{
				result = html2doct (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_HTMLZIP2DOCT:
			{
				result = html_zip2doct (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_HTML2DOCT_BIN:
			{
				result = html2doct_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;
			case TCD_HTMLZIP2DOCT_BIN:
			{
				result = html_zip2doct_bin (sFileFrom, sFileTo, sTempDir, oInputParams);
			}break;			
			//TCD_FB22DOCX,
			//TCD_FB22DOCT,
			//TCD_FB22DOCT_BIN,

			//TCD_EPUB2DOCX,
			//TCD_EPUB2DOCT,
			//TCD_EPUB2DOCT_BIN,
		}

		// delete temp dir
		if (!bExternalTempDir)
		{
			NSDirectory::DeleteDirectory(sTempDir);
		}

		//clean up v8
		// NSDoctRenderer::CDocBuilder::Dispose();
		if (SUCCEEDED_X2T(result) && oInputParams.m_bOutputConvertCorrupted)
		{
			return AVS_FILEUTILS_ERROR_CONVERT_CORRUPTED;
		}
		else
		{
			return result;
		}
	}

	_UINT32 FromFile(const std::wstring& file)
	{
		InputParams oInputParams;
		if(oInputParams.FromXmlFile(file))
		{
			return fromInputParams(oInputParams);
		}
		else
		{
			return AVS_FILEUTILS_ERROR_CONVERT_PARAMS;
		}
	}

	_UINT32 FromXml(const std::wstring& xml)
	{
		InputParams oInputParams;
		if(oInputParams.FromXml(xml))
		{
			return fromInputParams(oInputParams);
		}
		else
		{
			return AVS_FILEUTILS_ERROR_CONVERT_PARAMS;
		}
	}
}

