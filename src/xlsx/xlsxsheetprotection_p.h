/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
** All right reserved.
**
** Permission is hereby granted, free of charge, to any person obtaining
** a copy of this software and associated documentation files (the
** "Software"), to deal in the Software without restriction, including
** without limitation the rights to use, copy, modify, merge, publish,
** distribute, sublicense, and/or sell copies of the Software, and to
** permit persons to whom the Software is furnished to do so, subject to
** the following conditions:
**
** The above copyright notice and this permission notice shall be
** included in all copies or substantial portions of the Software.
**
** THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
** EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
** MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
** NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
** LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
** OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
** WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**
****************************************************************************/

#ifndef QXLSX_XLSXSHEETPROTECTION_P_H
#define QXLSX_XLSXSHEETPROTECTION_P_H

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include "xlsxsheetprotection.h"
#include <QSharedData>

QT_BEGIN_NAMESPACE_XLSX

class Q_XLSX_EXPORT  SheetProtectionPrivate : public QSharedData
{
public:
    SheetProtectionPrivate();
    SheetProtectionPrivate(bool enabled, bool useEditObjets, bool useEditScenarios, bool formatCells, bool formatColumns, bool formatRows, bool insertColumns, bool insertRows,
        bool insertHyperlinks, bool deleteColumns, bool deleteRows, bool sort, bool useAutoFilter, bool usePivotTablereports, bool selectLockedCells,
        bool selectUnLockedCells);
    SheetProtectionPrivate(const SheetProtectionPrivate &other);
    ~SheetProtectionPrivate();

    bool enabled;

    bool canSelectLockedCells;
    bool canSelectUnLockedCells;
    bool canFormatCells;
    bool canFormatColumns;
    bool canFormatRows;
    bool canInsertColumns;
    bool canInsertRows;
    bool canInsertHyperlinks;
    bool canDeleteColumns;
    bool canDeleteRows;
    bool canSort;
    bool canUseAutoFilter;
    bool canUsePivotTablereports;
    bool canUseEditObjects;
    bool canUseEditScenarios;
};

QT_END_NAMESPACE_XLSX
#endif // QXLSX_XLSXSHEETPROTECTION_P_H

