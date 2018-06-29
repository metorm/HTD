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
#include <QFileInfo>
#include <QAxObject>
#include <QAxWidget>
#include <QVariant>
#include <iostream>
#include <QThread>
#include <QApplication>

int getWDFormat(const QString & filePath)
{
    QFileInfo fileInfo(filePath);
    const QString ext = fileInfo.suffix();
    if(ext == QString("doc"))
        return 0;
    else if (ext == QString("rtf"))
        return 6;
    else if (ext.startsWith("htm"))
        return 8;
    else if (ext.startsWith("mhtm"))
        return 10;
    else
        return 0;
}

bool convert(const QString & in, const QString & out, bool landscape)
{
    std::cout << "INPUT: "
              << in.toStdString()
              << std::endl
              << "OUTPUT: "
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

    // ------- Save as --------
    activeDocument->dynamicCall(
                "SaveAs(const QVariant&, const QVariant&)",
                QVariant(out),
                QVariant(getWDFormat(out)));

    // unlink figures
    QAxObject* selection = wordAx.querySubObject("Selection");
    selection->dynamicCall("WholeStory()");
    QAxObject* fields = selection->querySubObject("Fields");
    fields->dynamicCall("Unlink()");
    activeDocument->dynamicCall("save()");

    // close & quit
    activeDocument->dynamicCall("Close (boolean)", false);
    wordAx.dynamicCall("Quit()");

    std::cout << "FINISHED" << std::endl;

    return true;
}
