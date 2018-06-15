/*
 * Copyright (C) 2015-2018 Ran Wei
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

#include <QString>
#include <QAxObject>
#include <QAxWidget>
#include <QVariant>
#include <iostream>
#include <QThread>
#include <QApplication>

bool convert(const QString & in, const QString & out, bool landscape)
{
    std::cout << "input: "
              << in.toStdString()
              << std::endl
              << "output: "
              << out.toStdString()
              << std::endl;

    const QString winWordAxName("Word.Application");
    QAxWidget wordAx(winWordAxName);
    if(wordAx.isNull())
    {
        std::cout << "FAIL: Cannot load COM object named \"" << winWordAxName.toStdString() << "\"" << std::endl;
    }
    wordAx.setProperty("Visible", false);

    QAxObject * wordDoc = wordAx.querySubObject("Documents");

    // ------- open input -------
    {
        QVariant filename(in);
        QVariant confirmconversions(false);
        QVariant readonly(false);
        QVariant addtorecentfiles(false);
        QVariant passworddocument("");
        QVariant passwordtemplate("");
        QVariant revert(false);
        wordDoc->querySubObject(
                    "Open(const QVariant&, const QVariant&,const QVariant&, const QVariant&, const QVariant&, const QVariant&,const QVariant&)",
                    filename,
                    confirmconversions,
                    readonly,
                    addtorecentfiles,
                    passworddocument,
                    passwordtemplate,
                    revert);
    }

    QAxObject* activeDocument = wordAx.querySubObject("ActiveDocument");

    // ------- set view -------
    QAxObject * activeWindow = wordAx.querySubObject("ActiveWindow");
    QAxObject * currentView = activeWindow->querySubObject("View");
    currentView->setProperty("Type", "wdPrintView");

    if(landscape)
    {
        QAxObject *pageSetup  = activeDocument->querySubObject("PageSetup");
        pageSetup->setProperty("Orientation", "wdOrientLandscape");
    }

    activeDocument->dynamicCall("SaveAs (const QString&)", out);
    activeDocument->dynamicCall("Close (boolean)", false);
    wordAx.dynamicCall("Quit()");

    std::cout << "FINISHED" << std::endl;

    return true;
}