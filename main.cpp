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

#include <QApplication>
#include <QCommandLineOption>
#include <QCommandLineParser>
#include <QString>
#include <QStringList>
#include <QFileInfo>
#include <QDir>

#include <iostream>

bool convert(const QString & in, const QString & out, bool landscape);

bool checkFileExtensions(const QString & path, const QStringList & ext)
{
    if(ext.size() == 0) return false;

    bool r = false;

    for(QStringList::const_iterator i = ext.constBegin(); i != ext.constEnd(); ++i)
    {
        if(path.endsWith(*i))
        {
            r = true;
            break;
        }
    }

    return r;
}

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);

    std::cout << "INFO: For latest version or other information, visit https://github.com/metorm/HTD" << std::endl;

    const QString appDescription("Html to Microsoft Word converter");
    QApplication::setApplicationName(appDescription);

    QCommandLineParser parser;
    parser.setApplicationDescription(appDescription);
    parser.addHelpOption();
    parser.addPositionalArgument("input", QCoreApplication::translate("main", "Full path to the HTML file (or any other parsable formats of MS Word). Only ASCII characters are allowed."));
    parser.addPositionalArgument("output", QCoreApplication::translate("main", "Full path to the output file. Only ASCII characters are allowed."));

    QCommandLineOption landscapeOption(QStringList() << "l" << "landscape",
                                       QCoreApplication::translate("main", "Set the paper orientation to landscape"));
    parser.addOption(landscapeOption);

    parser.process(app);

    const QStringList args = parser.positionalArguments();

    QStringList acceptableFileExtensions;
    // Add docx if you ensure this is run with Office 2007 or above
    // After adding any acceptable format, add corresponding tag in convert.cpp - getWDFormat
    acceptableFileExtensions << "doc" << "html" << "htm" << "mht" << "mhtml" << "txt" << "rtf";

    if(args.size() >= 2)
    {
        const QString inputFile = args.at(0);
        const QString outputFile = args.at(1);
        const QFileInfo inputFileInfo(inputFile);
        const QFileInfo outputFileInfo(outputFile);

        if(inputFileInfo.isDir() || outputFileInfo.isDir())
        {
            std::cout << "ERROR: Path to input or output file is a directory!" << std::endl;
            return -1;
        }

        if (!(checkFileExtensions(inputFile, acceptableFileExtensions) && checkFileExtensions(outputFile, acceptableFileExtensions)))
        {
            std::cout << "ERROR: Invalid input or output file format!" << std::endl;
            return -1;
        }

        if (!(inputFileInfo.isReadable()))
        {
            std::cout << "ERROR: Input file " << inputFile.toStdString() << " doesn't exist or is not readable!" << std::endl;
            return -1;
        }

        if (!(outputFileInfo.isWritable() || QFileInfo(outputFileInfo.dir().path()).isWritable()))
        {
            std::cout << "ERROR: Output file " << outputFile.toStdString() << " is not writable!" << std::endl;
            return -1;
        }

        if((inputFileInfo.absolutePath() + inputFileInfo.baseName()) ==
                (outputFileInfo.absolutePath() + outputFileInfo.baseName()))
        {
            std::cout << "FORBIDDEN: The difference between input and output path CANNOT only lies in the suffix!" << std::endl;
            std::cout << "FORBIDDEN: Choose a different path OR file name for the output file." << std::endl;
            return -1;
        }

        {
            const QString inputFilePathForConvert = QDir::toNativeSeparators(inputFileInfo.absoluteFilePath());
            const QString outputFilePathForConvert = QDir::toNativeSeparators(outputFileInfo.absoluteFilePath());

            return convert(inputFilePathForConvert,
                           outputFilePathForConvert,
                           parser.isSet(landscapeOption)) ?
                        0 : -1;
        }
    }else {
        std::cout << "ERROR: Paths to the input and output files are required!" << std::endl;
        return -1;
    }
}

