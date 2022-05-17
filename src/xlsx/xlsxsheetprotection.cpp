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

#include "xlsxworksheet.h"
#include "xlsxsheetprotection.h"
#include "xlsxsheetprotection_p.h"

#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_XLSX

SheetProtectionPrivate::SheetProtectionPrivate()
    :enabled(false), canSelectLockedCells(true), canSelectUnLockedCells(true), canFormatCells(false), canFormatColumns(false), canFormatRows(false)
    , canInsertColumns(false), canInsertRows(false), canInsertHyperlinks(false), canDeleteColumns(false), canDeleteRows(false), canSort(false)
    , canUseAutoFilter(false), canUsePivotTablereports(false), canUseEditObjects(true), canUseEditScenarios(true)
{

}

SheetProtectionPrivate::SheetProtectionPrivate(bool enabled, bool useEditObjets, bool useEditScenarios, bool formatCells, bool formatColumns, bool formatRows
    , bool insertColumns, bool insertRows, bool insertHyperlinks, bool deleteColumns, bool deleteRows, bool sort, bool useAutoFilter, bool usePivotTablereports
    , bool selectLockedCells, bool selectUnLockedCells)
    :enabled(enabled), canSelectLockedCells(selectLockedCells), canSelectUnLockedCells(selectUnLockedCells), canFormatCells(formatCells), canFormatColumns(formatColumns)
    , canFormatRows(formatRows), canInsertColumns(insertColumns), canInsertRows(insertRows), canInsertHyperlinks(insertHyperlinks), canDeleteColumns(deleteColumns)
    , canDeleteRows(deleteRows), canSort(sort), canUseAutoFilter(useAutoFilter), canUsePivotTablereports(usePivotTablereports)
    , canUseEditObjects(useEditObjets), canUseEditScenarios(useEditScenarios)
{

}

SheetProtectionPrivate::SheetProtectionPrivate(const SheetProtectionPrivate &other)
    :QSharedData(other)
{

}

SheetProtectionPrivate::~SheetProtectionPrivate()
{

}

SheetProtection::SheetProtection()
    :d(new SheetProtectionPrivate())
{

}


/*!
    Constructs a copy of \a other.
*/
SheetProtection::SheetProtection(const SheetProtection &other)
    :d(other.d)
{

}

SheetProtection::SheetProtection(bool enabled, bool useEditObjets, bool useEditScenarios, bool formatCells, bool formatColumns, bool formatRows, bool insertColumns
    , bool insertRows, bool insertHyperlinks, bool deleteColumns, bool deleteRows, bool sort, bool useAutoFilter, bool usePivotTablereports, bool selectLockedCells,
    bool selectUnLockedCells)
    :d(new SheetProtectionPrivate(enabled, useEditObjets, useEditScenarios, formatCells, formatColumns, formatRows, insertColumns, insertRows, insertHyperlinks,
        deleteColumns, deleteRows, sort, useAutoFilter, usePivotTablereports, selectLockedCells, selectUnLockedCells))
{

}

/*!
    Assigns \a other to this validation and returns a reference to this validation.
 */SheetProtection &SheetProtection::operator=(const SheetProtection &other)
{
    this->d = other.d;
    return *this;
}


/*!
 * Destroy the object.
 */
SheetProtection::~SheetProtection()
{
}

/*!
Returns the enabled.
*/
bool SheetProtection::enabled() const
{
    return d->enabled;
}

/*!
    Returns the selectLockedCells.
 */
bool SheetProtection::canSelectLockedCells() const
{
    return d->canSelectLockedCells;
}

/*!
Returns the selectUnLockedCells.
*/
bool SheetProtection::canSelectUnLockedCells() const
{
    return d->canSelectUnLockedCells;
}

/*!
Returns the formatCells.
*/
bool SheetProtection::canFormatCells() const
{
    return d->canFormatCells;
}

/*!
Returns the formatColumns.
*/
bool SheetProtection::canFormatColumns() const
{
    return d->canFormatColumns;
}

/*!
Returns the formatRows.
*/
bool SheetProtection::canFormatRows() const
{
    return d->canFormatRows;
}

/*!
Returns the insertColumns.
*/
bool SheetProtection::canInsertColumns() const
{
    return d->canInsertColumns;
}

/*!
Returns the insertRows.
*/
bool SheetProtection::canInsertRows() const
{
    return d->canInsertRows;
}

/*!
Returns the insertHyperlinks.
*/
bool SheetProtection::canInsertHyperlinks() const
{
    return d->canInsertHyperlinks;
}

/*!
Returns the deleteColumns.
*/
bool SheetProtection::canDeleteColumns() const
{
    return d->canDeleteColumns;
}

/*!
Returns the deleteRows.
*/
bool SheetProtection::canDeleteRows() const
{
    return d->canDeleteRows;
}

/*!
Returns the sort.
*/
bool SheetProtection::canSort() const
{
    return d->canSort;
}

/*!
Returns the useAutoFilter.
*/
bool SheetProtection::canUseAutoFilter() const
{
    return d->canUseAutoFilter;
}

/*!
Returns the usePivotTablereports.
*/
bool SheetProtection::canUsePivotTablereports() const
{
    return d->canUsePivotTablereports;
}

/*!
Returns the useEditObjects.
*/
bool SheetProtection::canUseEditObjects() const
{
    return d->canUseEditObjects;
}

/*!
Returns the useEditScenarios.
*/
bool SheetProtection::canUseEditScenarios() const
{
    return d->canUseEditScenarios;
}

/*!
Sets the setEnabled.
*/
void SheetProtection::setEnabled(bool enable)
{
    d->enabled = enable;
}

/*!
Sets the selectLockedCells.
*/
void SheetProtection::allowUserToSelectLockedCells(bool enable)
{
    d->canSelectLockedCells = enable;
}

/*!
Sets the selectUnLockedCells.
*/
void SheetProtection::allowUserToSelectUnLockedCells(bool enable)
{
    d->canSelectUnLockedCells = enable;
}

/*!
Sets the formatCells.
*/
void SheetProtection::allowUserToFormatCells(bool enable)
{
    d->canFormatCells = enable;
}

/*!
Sets the selectUnLockedCells.
*/
void SheetProtection::allowUserToFormatColumns(bool enable)
{
    d->canFormatColumns = enable;
}

/*!
Sets the selectUnLockedCells.
*/
void SheetProtection::allowUserToFormatRows(bool enable)
{
    d->canFormatRows = enable;
}

/*!
Sets the insertColumns.
*/
void SheetProtection::allowUserToInsertColumns(bool enable)
{
    d->canInsertColumns = enable;
}

/*!
Sets the insertRows.
*/
void SheetProtection::allowUserToInsertRows(bool enable)
{
    d->canInsertRows = enable;
}

/*!
Sets the insertHyperlinks.
*/
void SheetProtection::allowUserToInsertHyperlinks(bool enable)
{
    d->canInsertHyperlinks = enable;
}

/*!
Sets the deleteColumns.
*/
void SheetProtection::allowUserToDeleteColumns(bool enable)
{
    d->canDeleteColumns = enable;
}

/*!
Sets the deleteRows.
*/
void SheetProtection::allowUserToDeleteRows(bool enable)
{
    d->canDeleteRows = enable;
}

/*!
Sets the sort.
*/
void SheetProtection::allowUserToSort(bool enable)
{
    d->canSort = enable;
}

/*!
Sets the useAutoFilter.
*/
void SheetProtection::allowUserToUseAutoFilter(bool enable)
{
    d->canUseAutoFilter = enable;
}

/*!
Sets the usePivotTablereports.
*/
void SheetProtection::allowUserToUsePivotTablereports(bool enable)
{
    d->canUsePivotTablereports = enable;
}

/*!
Sets the useEditObjects.
*/
void SheetProtection::allowUserToUseEditObjects(bool enable)
{
    d->canUseEditObjects = enable;
}

/*!
Sets the useEditScenarios.
*/
void SheetProtection::allowUserToUseEditScenarios(bool enable)
{
    d->canUseEditScenarios = enable;
}

/*!
 * \internal
 */
bool SheetProtection::saveToXml(QXmlStreamWriter &writer) const
{
    if (enabled()) {

        writer.writeAttribute(QStringLiteral("sheet"), QStringLiteral("1"));

        if (!canUseEditObjects()) {
            writer.writeAttribute(QStringLiteral("objects"), QStringLiteral("1"));
        }

        if (!canUseEditScenarios()) {
            writer.writeAttribute(QStringLiteral("scenarios"), QStringLiteral("1"));
        }

        if (canFormatCells()) {
            writer.writeAttribute(QStringLiteral("formatCells"), QStringLiteral("0"));
        }

        if (canFormatColumns()) {
            writer.writeAttribute(QStringLiteral("formatColumns"), QStringLiteral("0"));
        }

        if (canFormatRows()) {
            writer.writeAttribute(QStringLiteral("formatRows"), QStringLiteral("0"));
        }

        if (canInsertColumns()) {
            writer.writeAttribute(QStringLiteral("insertColumns"), QStringLiteral("0"));
        }

        if (canInsertRows()) {
            writer.writeAttribute(QStringLiteral("insertRows"), QStringLiteral("0"));
        }

        if (canInsertHyperlinks()) {
            writer.writeAttribute(QStringLiteral("insertHyperlinks"), QStringLiteral("0"));
        }

        if (canDeleteColumns()) {
            writer.writeAttribute(QStringLiteral("deleteColumns"), QStringLiteral("0"));
        }

        if (canDeleteRows()) {
            writer.writeAttribute(QStringLiteral("deleteRows"), QStringLiteral("0"));
        }

        if (canSort()) {
            writer.writeAttribute(QStringLiteral("sort"), QStringLiteral("0"));
        }

        if (canUseAutoFilter()) {
            writer.writeAttribute(QStringLiteral("autoFilter"), QStringLiteral("0"));
        }

        if (canUsePivotTablereports()) {
            writer.writeAttribute(QStringLiteral("pivotTables"), QStringLiteral("0"));
        }

        if (!canSelectLockedCells() && !canSelectUnLockedCells()) {
            writer.writeAttribute(QStringLiteral("selectLockedCells"), QStringLiteral("1"));
            writer.writeAttribute(QStringLiteral("selectUnlockedCells"), QStringLiteral("1"));
        } else if (!canSelectLockedCells()) {
            writer.writeAttribute(QStringLiteral("selectLockedCells"), QStringLiteral("1"));
        }
    }

    return true;
}

/*!
 * \internal
 */
SheetProtection SheetProtection::loadFromXml(QXmlStreamReader &reader)
{
    SheetProtection sheetProtection;

    QXmlStreamAttributes attrs = reader.attributes();

    if (attrs.hasAttribute(QLatin1String("sheet"))) {
        sheetProtection.setEnabled((bool) attrs.value(QLatin1String("sheet")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("objects"))) {
        sheetProtection.allowUserToUseEditObjects(!(bool)attrs.value(QLatin1String("objects")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("scenarios"))) {
        sheetProtection.allowUserToUseEditScenarios(!(bool)attrs.value(QLatin1String("scenarios")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("formatCells"))) {
        sheetProtection.allowUserToFormatCells(!(bool)attrs.value(QLatin1String("formatCells")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("formatColumns"))) {
        sheetProtection.allowUserToFormatColumns(!(bool)attrs.value(QLatin1String("formatColumns")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("formatRows"))) {
        sheetProtection.allowUserToFormatRows(!(bool)attrs.value(QLatin1String("formatRows")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("insertColumns"))) {
        sheetProtection.allowUserToInsertColumns(!(bool)attrs.value(QLatin1String("insertColumns")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("insertRows"))) {
        sheetProtection.allowUserToInsertRows(!(bool)attrs.value(QLatin1String("insertRows")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("insertHyperlinks"))) {
        sheetProtection.allowUserToInsertHyperlinks(!(bool)attrs.value(QLatin1String("insertHyperlinks")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("deleteColumns"))) {
        sheetProtection.allowUserToDeleteColumns(!(bool)attrs.value(QLatin1String("deleteColumns")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("deleteRows"))) {
        sheetProtection.allowUserToDeleteRows(!(bool)attrs.value(QLatin1String("deleteRows")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("sort"))) {
        sheetProtection.allowUserToSort(!(bool)attrs.value(QLatin1String("sort")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("autoFilter"))) {
        sheetProtection.allowUserToUseAutoFilter(!(bool)attrs.value(QLatin1String("autoFilter")).toInt());
    }

    if (attrs.hasAttribute(QLatin1String("pivotTables"))) {
        sheetProtection.allowUserToUsePivotTablereports(!(bool)attrs.value(QLatin1String("pivotTables")).toInt());
    }

    bool userCanSelectLockedCells	= !(bool)attrs.value(QLatin1String("selectLockedCells")).toInt();
    bool userCanSelectUnLockedCells = !(bool)attrs.value(QLatin1String("selectUnlockedCells")).toInt();

    if (userCanSelectLockedCells) {
        sheetProtection.allowUserToSelectLockedCells(true);
        sheetProtection.allowUserToSelectUnLockedCells(true);
    }
    else if (userCanSelectUnLockedCells) {
        sheetProtection.allowUserToSelectLockedCells(false);
        sheetProtection.allowUserToSelectUnLockedCells(true);
    }
    else {
        sheetProtection.allowUserToSelectLockedCells(false);
        sheetProtection.allowUserToSelectUnLockedCells(false);
    }

    return sheetProtection;
}

QT_END_NAMESPACE_XLSX
