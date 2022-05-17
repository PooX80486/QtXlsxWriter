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
#ifndef QXLSX_XLSXSHEETPROTECTION_H
#define QXLSX_XLSXSHEETPROTECTION_H

#include "xlsxglobal.h"
#include <QSharedDataPointer>

class QXmlStreamReader;
class QXmlStreamWriter;

QT_BEGIN_NAMESPACE_XLSX

class Worksheet;
class SheetProtectionPrivate;

class Q_XLSX_EXPORT SheetProtection
{
public:

    SheetProtection();
    SheetProtection(const SheetProtection &other);
    SheetProtection(bool enabled, bool useEditObjets, bool useEditScenarios, bool formatCells, bool formatColumns, bool formatRows, bool insertColumns, bool insertRows,
        bool insertHyperlinks, bool deleteColumns, bool deleteRows, bool sort, bool useAutoFilter, bool usePivotTablereports, bool selectLockedCells,
        bool selectUnLockedCells);
    ~SheetProtection();

    bool enabled() const;
    void setEnabled(bool enable);

    bool canSelectLockedCells() const;
    bool canSelectUnLockedCells() const;
    bool canFormatCells() const;
    bool canFormatColumns() const;
    bool canFormatRows() const;
    bool canInsertColumns() const;
    bool canInsertRows() const;
    bool canInsertHyperlinks() const;
    bool canDeleteColumns() const;
    bool canDeleteRows() const;
    bool canSort() const;
    bool canUseAutoFilter() const;
    bool canUsePivotTablereports() const;
    bool canUseEditObjects() const;
    bool canUseEditScenarios() const;

    void allowUserToSelectLockedCells(bool enable);
    void allowUserToSelectUnLockedCells(bool enable);
    void allowUserToFormatCells(bool enable);
    void allowUserToFormatColumns(bool enable);
    void allowUserToFormatRows(bool enable);
    void allowUserToInsertColumns(bool enable);
    void allowUserToInsertRows(bool enable);
    void allowUserToInsertHyperlinks(bool enable);
    void allowUserToDeleteColumns(bool enable);
    void allowUserToDeleteRows(bool enable);
    void allowUserToSort(bool enable);
    void allowUserToUseAutoFilter(bool enable);
    void allowUserToUsePivotTablereports(bool enable);
    void allowUserToUseEditObjects(bool enable);
    void allowUserToUseEditScenarios(bool enable);

    SheetProtection &operator=(const SheetProtection &other);

    bool saveToXml(QXmlStreamWriter &writer) const;
    static SheetProtection loadFromXml(QXmlStreamReader &reader);
private:
    QSharedDataPointer<SheetProtectionPrivate> d;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXSHEETPROTECTION_H
