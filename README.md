# HTD
Portable Html to Microsoft Word converter

This tiny portable tool is designed to convert html pages to Microsoft word format.
But it should also applies to other acceptable formats such as `*rft`, `*.txt`, etc.

# Usage

```
C:\Documents\dev\build-HTD-Static-Release\release>HTD.exe --help

INFO: For latest version or other information, visit https://github.com/metorm/HTD
Usage: HTD.exe [options] input output
Html to Microsoft Word converter

Options:
  -?, -h, --help   Displays this help.
  -l, --landscape  Set the paper orientation to landscape

Arguments:
  input            Full path to the HTML file (or any other parsable formats of
                   MS Word). Only ASCII characters are allowed.
  output           Full path to the output file. Only ASCII characters are
                   allowed.

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

## Convert other formats

Just pass file paths having different suffixes:

`HTD.exe C:\Documents\webpage\111.doc C:\Documents\webpage\_111.rtf`


## Setting the orientation of the output file

Sometimes you may have very wide contents in your html file. In case ordinary A4 paper is not wide enough to hold your page, pass a `-l, --landscape` option to set the paper orientation to landscape.

# Build

This a tiny project. Just download/clone the source code and build it using `Qt Creator`.
There is no dependency other than `QT` itself. But you must have chosen `qtactiveqt` module when installing or building `QT`.

Or, you can use the pre-built binary file [HERE](https://github.com/metorm/HTD/releases). The released x86 WIN-XP-compatable binary is compiled with `QT` linked statically.

# Feel free to post an issue if you have any further problem.

----------------------------------------

# HTD

HTD是一个绿色版小工具，用于将 Html 文档转换到 Microsoft Word 兼容格式。

理论上也应该适用于其它 Word 软件可接受的格式，例如* rft，* .txt等。


# 使用

帮助信息：
```
INFO: For latest version or other information, visit https://github.com/metorm/HTD
Usage: HTD.exe [options] input output
Html to Microsoft Word converter

Options:
  -?, -h, --help   Displays this help.
  -l, --landscape  Set the paper orientation to landscape

Arguments:
  input            Full path to the HTML file (or any other parsable formats of
                   MS Word). Only ASCII characters are allowed.
  output           Full path to the output file. Only ASCII characters are
                   allowed.

```

## 将 `*.html` 转换到 `*.doc`

操作十分简单：
```
C:\Documents\dev\build-HTD-Static-Release\release>HTD.exe C:\Documents\webpage\111.html C:\Documents\webpage\111_.doc
input: C:\Documents\webpage\111.html
output: C:\Documents\webpage\111_.doc
FINISHED
```

**注意:** 输入/输出路径**不能**仅有后缀名不同. 例如: `HTD.exe C:\Documents\webpage\111.html C:\Documents\webpage\111.doc` 将使得 Word 引擎抛出异常, 而`HTD.exe C:\Documents\webpage\111.html C:\Documents\webpage\_111.doc`则可以正常工作.

## 转换其他格式

直接在输入输出路径中传递你想要的后缀名即可，例如：

`HTD.exe C:\Documents\webpage\111.doc C:\Documents\webpage\_111.rtf`

## 设置输出文件的纸张方向

有时候你可能在你的html文件中有很宽的内容。 如果普通A4纸张的宽度不足以容纳您的页面，请附加`-l，--landscape`选项，以将纸张方向设置为横向。

# 编译

本程序很小，编译十分简单。 只需下载/克隆源代码并使用`Qt Creator`打开、编译即可。
除了`QT`本身之外，没有任何依赖。 但是在安装或自行编译`QT`时，必须选择`qtactiveqt`模块。

如果没有开发需求，建议直接使用[预编译的二进制文件](https://github.com/metorm/HTD/releases)。所发布的文件兼容于Windows XP或以上版本，并静态链接了QT运行库。

# 如有问题或功能需求，欢迎发issue
