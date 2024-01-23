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

#include "cextracttools.h"
#include "ASCConverters.h"

#include <iostream>
#include <assert.h>
#include <stdio.h>
#include <string.h>
#include <emscripten.h>


#ifdef NODERAWFS
#define CWD ""
#else
#define CWD "/working/"
#endif

#define VALUE_TO_STRING(x) #x
#define VALUE(x) VALUE_TO_STRING(x)

using namespace NExtractTools;

#if !defined(_WIN32) && !defined (_WIN64)
static std::wstring utf8_to_unicode(const char *src)
{
    if (src == NULL) return _T("");
    std::string temp = src;

    unsigned int nLength = temp.length();

    UTF32 *pStrUtf32 = new UTF32 [nLength+1];
    memset ((void *) pStrUtf32, 0, sizeof (UTF32) * (nLength+1));


    UTF8 *pStrUtf8 = (UTF8 *) src;

    // this values will be modificated
    const UTF8 *pStrUtf8_Conv = pStrUtf8;
    UTF32 *pStrUtf32_Conv = pStrUtf32;

    ConversionResult eUnicodeConversionResult = ConvertUTF8toUTF32 (&pStrUtf8_Conv,
                                                                    &pStrUtf8[nLength]
                                                                    , &pStrUtf32_Conv
                                                                    , &pStrUtf32 [nLength]
                                                                    , strictConversion);

    if (conversionOK != eUnicodeConversionResult)
    {
        delete [] pStrUtf32;
        return L"";
    }
    std::wstring wsEntryName ((wchar_t *) pStrUtf32);

    delete [] pStrUtf32;
    return wsEntryName;
}
#endif


extern "C" {
int EMSCRIPTEN_KEEPALIVE runX2T(char *arg)
{
  FILE *file;
  int res;
  char buffer[512];

	std::wstring sArg1, sArg2, sExePath;

        std::cout << "Starting x2t" << std::endl;
#if !defined(_WIN32) && !defined (_WIN64)
    sArg1       = utf8_to_unicode(arg);
        std::wcout << "Param file: " << sArg1 << std::endl;
#else
    sArg1		= std::wstring(arg);
#endif

	_UINT32 result = 0;
    std::wstring sXmlExt = _T(".xml");
    if((sArg1.length() > 3) && (sXmlExt == sArg1.substr(sArg1.length() - sXmlExt.length(), sXmlExt.length())))
	{
        std::cout << "Reading param file" << std::endl;
		NExtractTools::InputParams oInputParams;
		if (oInputParams.FromXmlFile(sArg1) && (sArg2.empty() || oInputParams.FromXml(sArg2)))
		{
        std::wcout << "From file " << *oInputParams.m_sFileFrom << " To file " << *oInputParams.m_sFileTo << std::endl;
        std::cout << "Before conversion" << std::endl;
			result = NExtractTools::fromInputParams(oInputParams);
        std::cout << "After conversion" << std::endl;
		}
		else
		{
			result = AVS_FILEUTILS_ERROR_CONVERT_PARAMS;
		}
	}
	return getReturnErrorCode(result);
}
}

#if !defined(_WIN32) && !defined (_WIN64)
	int EMSCRIPTEN_KEEPALIVE main(int argc, char *argv[])
#else
	int wmain(int argc, wchar_t *argv[])
#endif
{
  if (argc>1) {
   // mount the current folder as a NODEFS instance
   // inside of emscripten
   EM_ASM(
    FS.mkdir('/working');
    FS.mount(NODEFS, { root: '.' }, '/working');
   );
   return runX2T(argv[1]);
   }
  else
   return 0;
}
