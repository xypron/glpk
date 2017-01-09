/* dll.c (dll entry point) */

/***********************************************************************
*  This code is part of GLPK (GNU Linear Programming Kit).
*
*  Author Heinrich Schuchardt <xypron.glpk@gmx.de>
*
*  Copyright (C) 2017 Andrew Makhorin, Department for Applied
*  Informatics, Moscow Aviation Institute, Moscow, Russia. All rights
*  reserved. E-mail: <mao@gnu.org>.
*
*  GLPK is free software: you can redistribute it and/or modify it
*  under the terms of the GNU General Public License as published by
*  the Free Software Foundation, either version 3 of the License, or
*  (at your option) any later version.
*
*  GLPK is distributed in the hope that it will be useful, but WITHOUT
*  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
*  or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public
*  License for more details.
*
*  You should have received a copy of the GNU General Public License
*  along with GLPK. If not, see <http://www.gnu.org/licenses/>.
***********************************************************************/


#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#ifdef __WOE__

#pragma comment(lib, "user32.lib")
#include <windows.h>

#define VISTA 0x06

/*
 * This is the main entry point of the DLL.
 */
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{  DWORD version;
   DWORD major_version;

#ifdef TLS
   switch (fdwReason) {
      case DLL_PROCESS_ATTACH:
         /*
          * @TODO:
          * GetVersion is deprecated but the version help functions are not
          * available in Visual Studio 2010. So lets use it until we remove
          * the outdated Build files.
          */
         version = GetVersion();
         major_version = version & 0xff;

         if (major_version < VISTA) {
            MessageBoxA(NULL, "This GLPK library uses thread local storage.\n"
               "You need at least Windows Vista to utilize it.\n",
               "GLPK", MB_OK | MB_ICONERROR);
            return FALSE;
         }
         break;
   }
#endif /* TLS */

   return TRUE;  
}

#endif /* WOE */