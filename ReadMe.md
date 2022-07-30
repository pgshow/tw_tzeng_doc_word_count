## doc_word_count
可以批量计算 Word Document 的字数并输出每个 doc 的情况，目前的算法中文比较准确，英文多就会有误差。

### 使用方法
* Word 文档放入 Article 文件夹
* 批量检查结果会输出到 result.txt
* 运行时请不要操作 WPS，以免报错

```
编译成 exe
$ pyinstaller -F main.py
```

### 注意：
如果报错 **COMObject unknown**, 可能是 win32com 找不到 office 的 Word.Application 这个名称，此时可以安装 WPS 来代替。

#### cmd 窗口总是卡住解决方案：
https://blog.csdn.net/weixin_39858881/article/details/106935616