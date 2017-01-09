/* tls.c (thread local storage) */

/***********************************************************************
*  This code is part of GLPK (GNU Linear Programming Kit).
*
*  Copyright (C) 2001-2013 Andrew Makhorin, Department for Applied
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
#include <config.h>
#endif

#include "env.h"

#ifndef TLS
static void *tls = NULL;
#else
static TLS void *tls = NULL;
/* this option allows running multiple independent instances of GLPK in
 * different threads of a multi-threaded application, in which case the
 * variable tls should be placed in the Thread Local Storage (TLS);
 * it is assumed that the macro TLS is previously defined to something
 * like '__thread', '_Thread_local', etc. */
#endif

/***********************************************************************
*  NAME
*
*  tls_set_ptr - store global pointer in TLS
*
*  SYNOPSIS
*
*  #include "env.h"
*  void tls_set_ptr(void *ptr);
*
*  DESCRIPTION
*
*  The routine tls_set_ptr stores a pointer specified by the parameter
*  ptr in the Thread Local Storage (TLS). */

void tls_set_ptr(void *ptr)
{     tls = ptr;
      return;
}

/***********************************************************************
*  NAME
*
*  tls_get_ptr - retrieve global pointer from TLS
*
*  SYNOPSIS
*
*  #include "env.h"
*  void *tls_get_ptr(void);
*
*  RETURNS
*
*  The routine tls_get_ptr returns a pointer previously stored by the
*  routine tls_set_ptr. If the latter has not been called yet, NULL is
*  returned. */

void *tls_get_ptr(void)
{     void *ptr;
      ptr = tls;
      return ptr;
}

/* eof */
