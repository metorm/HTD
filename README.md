# HTD
Portable Html to Microsoft Word converter

This tiny portable tool is designed to convert html pages to Microsoft word format.
But it should also applies to other acceptable formats such as `*rft`, `*.txt`, etc.

# Usage

```
C:\Documents\dev\build-HTD-Static-Release\release>HTD.exe --help
Usage: HTD.exe [options] input output
Html to Microsoft Word converter

Options:
  -?, -h, --help   Displays this help.
  -l, --landscape  Set the paper orientation to landscape

Arguments:
  input            Path to the HTML file (or any other parsable formats of MS
                   Word)
  output           Path to the output file.

```

## Convert `*.html` to `*.doc`

The operation is quite straightforward:
```
C:\Documents\dev\build-HTD-Static-Release\release>HTD.exe C:\Documents\webpage\111.html C:\Documents\webpage\111_.doc
input: C:\Documents\webpage\111.html
output: C:\Documents\webpage\111_.doc
FINISHED
```

**Note:** The difference between input and output path CANNOT only lies in the suffix. E.g.: `HTD.exe C:\Documents\webpage\111.html C:\Documents\webpage\111.doc` causes the `Word` engine to throw a exception, while `HTD.exe C:\Documents\webpage\111.html C:\Documents\webpage\_111.doc` works well.

## Setting the orientation of the output file

Sometimes you may have very wide contents in your html file. In case ordinary A4 paper is not wide enough to hold your page, pass a `-l, --landscape` option to set the paper orientation to landscape.

# Build

This a tiny project. Just download/clone the source code and build it using `Qt Creator`.
There is no dependency other than `QT` itself. But you must have chosen `qtactiveqt` module when installing or building `QT`.

Or, you can use the pre-built binary file [HERE](https://github.com/metorm/HTD/releases). The released x86 WIN-XP-compatable binary is compiled with `QT` linked statically.

# Feel free to post an issue if you have any further problem.
