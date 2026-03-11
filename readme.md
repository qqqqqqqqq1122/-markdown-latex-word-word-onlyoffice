这绝对是最一劳永逸的做法！花这 10 分钟折腾完，你的 ONLYOFFICE 顶部就会永远固定一个“转公式”的专属按钮。写论文时，选中一按，直接变身。

这其实就是开发一个**“无界面的后台运行插件”**。一共分四步，非常简单：

### 第一步：建个“毛坯房”

1. 在你的电脑上（比如桌面或者 G 盘），新建一个普通的文件夹，名字随便起，比如叫 `LaTeX_Plugin`。
2. 随便找一张**正方形的小图片**（比如你平时存的任何小图标），把它复制到这个文件夹里，并重命名为 **`icon.png`**。（这是你未来要在工具栏上看到的按钮图标）。

### 第二步：给插件写个“身份证” (config.json)

1. 在 `LaTeX_Plugin` 文件夹里，右键 -> 新建 -> **文本文档**。
2. 把这个文本文档重命名为 **`config.json`**（注意：一定要删掉 `.txt` 后缀，如果你的电脑隐藏了后缀名，记得先在文件夹的“查看”里勾选“文件扩展名”）。
3. 用“记事本”打开这个 `config.json`，把下面这段配置代码原封不动地粘贴进去，然后保存并关闭：

```json
{
  "name": "转公式",
  "guid": "asc.{6B36ED69-6B18-40F0-91AF-5C2849845311}",
  "version": "1.0",
  "variations": [
    {
      "description": "将选中的LaTeX转化为原生公式",
      "url": "index.html",
      "icons": ["icon.png"],
      "isViewer": false,
      "EditorsSupport": ["word"],
      "isVisual": false,
      "isModal": false,
      "isInsideMode": false,
      "initDataType": "none",
      "initData": "",
      "buttons": []
    }
  ]
}

```

*(这里面的 `isVisual: false` 是神来之笔，意思是你点这个按钮时，它不会弹出烦人的对话框，而是在后台瞬间执行代码。)*

### 第三步：装入“发动机” (index.html)

1. 同样在这个文件夹里，再新建一个文本文档，重命名为 **`index.html`**。
2. 用记事本打开它，把你刚才在宏里测试通过的“脱壳转公式”代码，套进插件专用的格式里。直接复制下面这段完整代码粘贴进去，保存并关闭：

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <script type="text/javascript" src="https://onlyoffice.github.io/sdkjs-plugins/v1/plugins.js"></script>
    <script>
        // 插件初始化（当你点击按钮时触发）
        window.Asc.plugin.init = function() {
            
            // callCommand 用来将代码注入到文档底层运行
            this.callCommand(function() {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();
                if (!range) return;
                
                var text = range.GetText();
                if (!text || text.trim() === "") return;
                
                // 核心净水器
                var cleanText = text.trim()
                    .replace(/^(\$\$?|\\\[|\\\()\s*/, "")
                    .replace(/\s*(\$\$?|\\\]|\\\))$/, "");
                
                // 删除原文本并插入公式
                range.Delete();
                doc.AddMathEquation(cleanText, "latex");
                
            }, false, true, function() {
                // 代码运行完后，立刻自我关闭，不留痕迹
                window.Asc.plugin.executeCommand("close", "");
            });
        };
    </script>
</head>
<body style="margin: 0;"></body>
</html>

```

### 第四步：在 ONLYOFFICE 里“上牌照”

现在的 `LaTeX_Plugin` 文件夹里应该只有三个文件：`icon.png`、`config.json` 和 `index.html`。

1. 打开 ONLYOFFICE Desktop Editors。
2. 随便新建或打开一个 Word 文档，点击顶部菜单栏的 **“插件 (Plugins)”**。
3. 点击工具栏里的 **“管理插件 (Plugin Manager)”** 图标（通常是一个带齿轮或者积木的图标）。
4. 在弹出的窗口左侧，点击底部的 **“设置 (Settings)”**。
5. 点击 **“添加插件 (Add plugin)”** 按钮。
6. 在弹出的系统文件选择框里，**找到并选中你刚才写的那个 `config.json` 文件**，点击打开。

**大功告成！** 关掉管理窗口，看一眼你的“插件”选项卡顶部，是不是已经多了一个带有你那张图片的**“转公式”**按钮了？

**去文档里敲一段带 `$$` 的矩阵公式代码，选中它，直接去顶部点一下这个按钮，享受属于你自己的极客级“光标同步”体验吧！成功了吗？**


哎呀，这又是 ONLYOFFICE 给你挖的一个坑！

你看到的那个 **“提交您自己的插件”** 是给全球开发者用的，点进去会跳转到 GitHub 让你给官方提代码库，让全世界的人都能下载你的插件。咱们只是想自己在本地爽一下，根本不需要走这种繁琐的官方审核流程。

在 ONLYOFFICE 的最新版本里，它把本地安装按钮藏得很深（有些版本甚至在界面上移除了本地安装入口）。但这根本难不倒我们，因为我们可以直接进行**“底层文件夹暴力注入”**。

这比在界面上点来点去还要快，跟着我走这 3 步，1 分钟搞定：

### 第一步：给插件披上“系统身份证”

ONLYOFFICE 的底层只认一种文件夹名字，那就是一串代码。
把你刚刚建好的那个 `LaTeX_Plugin` 文件夹，**重命名**为这串精确的代码（连带大括号一起复制）：
**`{6B36ED69-6B18-40F0-91AF-5C2849845311}`**

*(请确保点开这个文件夹，里面直接就是 `config.json`、`index.html` 和 `icon.png` 这三个文件，没有多余的子文件夹)*

### 第二步：直达 ONLYOFFICE 的“神经中枢”

1. 按下键盘上的 **`Win + R`** 键，打开运行窗口。
2. 把下面这串系统路径原封不动地复制，粘贴进去，然后按回车：
`%UserProfile%\AppData\Local\ONLYOFFICE\DesktopEditors\data\sdkjs-plugins`
3. 这时会弹出一个文件夹，这里面装的就是 ONLYOFFICE 所有的底层插件。

### 第三步：注入并重启

1. 把你第一步重命名好的 **`{6B36ED69-6B18-40F0-91AF-5C2849845311}`** 整个文件夹，直接**拖动或粘贴**到这个 `sdkjs-plugins` 目录里。
2. **彻底关闭 ONLYOFFICE**（确保右下角系统托盘里也没有它在后台运行）。
3. 重新打开 ONLYOFFICE，随便进一个 Word 文档。

**现在，点开顶部的“插件”选项卡。** 不用进管理器了，你的专属“转公式”按钮应该已经高高地挂在工具栏上了！你去随便敲一段带 `$$` 的代码，选中它，点一下这个按钮，看看这感觉是不是终于达到了完美的“一键变身”？


##############################3
对于word，直接用宏导入这个文件就可以，AMCEqnNum.bas


太棒了！最难的底层逻辑终于彻底跑通了。现在我们把它“实体化”，变成你专属的工具栏按钮，以后写论文点一下就瞬间变身。

既然你引用了这些步骤，说明你已经准备好把宏放到最顺手的位置了。如果你在操作时找不到具体的按钮，请按我下面这个**更详细的“保姆级”放大版步骤**来点：

### 第一步：打开“自定义”大门

1. 抬头看你的 Word 窗口**最左上角**。那里通常有“自动保存”开关、一个软盘图标（保存）和一个撤销的弯箭头。
2. 在这一排小图标的最右边，有一个**极其微小的“向下指的横线+小三角”**图标（鼠标悬停会显示“自定义快速访问工具栏”）。
3. 点开它，在弹出的长菜单里，一直往下拉，找到并点击 **“其他命令... (More Commands)”**。

### 第二步：把宏“揪”出来

1. 这时会弹出一个大窗口（Word 选项）。
2. 注意看左上方有一个下拉菜单，默认写着“常用命令”。点开它，往下滚，选择 **“宏 (Macros)”**。
3. 选完之后，下方的大列表里就会出现你电脑里所有的宏。你要找的那个名字应该类似于：
`Normal.AMCEqnNum.LaTeX转公式` （也可能是 `Normal.NewMacros.LaTeX转公式`，认准后面这几个字就行）。
4. 鼠标**单击选中**它。
5. 点击两个大白框中间那个长方形的 **“添加 >> (Add)”** 按钮。你会看到这个宏跑到右边的大框里了。

### 第三步：给它做个“整容”（关键步骤，不然图标很丑）

如果现在直接点确定，它在工具栏上会是一个像流程图一样极其难看的默认图标，而且鼠标放上去会显示一长串代码名。我们来美化一下：

1. 在**右侧的框里**，选中你刚添加过去的那个宏。
2. 点击大白框正下方的 **“修改... (Modify...)”** 按钮。
3. 弹出的窗口里，你可以做两件事：
* **选图标：** 在那一堆密密麻麻的图标里，挑一个你顺眼的！我强烈建议选那个 **“π”** 符号，或者一个**笑脸**，或者一个**小闪电**，非常有专属定制的感觉。
* **改名字：** 在下方的“显示名称”框里，把那串长长的英文删掉，只保留 **`LaTeX转公式`** 这几个字。


4. 点“确定”关掉小窗口，再点“确定”关掉大窗口。

### 见证奇迹的时刻

现在，回到你的 Word 正文界面。抬头看最左上角，是不是已经多了一个你刚刚亲手挑选的小图标（比如 π）？

**你现在去文档里选中一段带 `$$` 的代码，然后鼠标直接去左上角点一下那个小图标。是不是瞬间就变成完美公式了？** 以后这台电脑上的 Word，就有了你亲手打造的“一键变身”功能！

#####################################

…or create a new repository on the command line
echo "# -markdown-latex-word-word-onlyoffice" >> README.md
git init
git add README.md
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/qqqqqqqqq1122/-markdown-latex-word-word-onlyoffice.git
git push -u origin main
…or push an existing repository from the command line
git remote add origin https://github.com/qqqqqqqqq1122/-markdown-latex-word-word-onlyoffice.git
git branch -M main
git push -u origin main

