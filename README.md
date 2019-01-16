# xmindToTestcase
Convert Xmind file to Excel testcases file.

## usage

To see ``xmindToTestcase`` version:

```shell
$ python3 start.py -V
0.0.1
```

## examples
you can run ``xmindToTestcase`` like this:
```shell
$ python3 start.py test/Demo.xmind test/Demo.xls
INFO:root:Generate Xmind file successfully: test/Demo.xls
```
As you see, the first parameter is xmind source file path, and the second is converted Excel file path.

## xmind examples
```shell
功能模块名称 - 功能名称 - 子功能名称 - 功能点 - 检查点 - 用例步骤 - 预期结果 - 预期的结果信息

```
