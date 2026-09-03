
【Git版本管理建立命令】
git init
git status
git status --short --branch
git diff
git add .
git commit -m "Update 20260524 - v0.6"
git log --oneline --graph --decorate


【执行命令】


【通用工作注意事项如下】

脚本成熟后，完整运行一次，诊断过程和输出是否完全正确，如果有错误就逐步排查修复，如果一切正常，就继续写一个中文Readme文档，放在脚本所在的目录中。Readme文档的开头就写出脚本的使用方式和参数的意义，便于打开后快速使用，之后再写出其他内容。Readme文档要表明这个脚本的依赖，并给出依赖安装的命令，再生成一个 requirements.txt，以便移植以后快速安装环境。顺便检查脚本具有自动检查依赖和环境的能力，如果移植后的环境中缺乏某些必须的内容或者报错，让脚本提醒使用者所需内容或解决方案。

建立git版本控制体系，每次修改的时候更新git并自动commit进去。

#调整文件夹结构和代码中的路径，并移动对应的文件去文件夹，归类规则如下：
#文件夹结构的要求如下，未来开发时注意文件归类规则如下：
1. 脚本位于项目文件夹根目录。你可以调用git来管理版本，如果有需要。
2. 子脚本位于“bin”目录。
3. 临时文件位于“temp”目录，“__pycache__”可以留在根目录，不需要去“temp”目录。
4. 输入文件位于“input”标识的目录，具体文件名可以根据你的建议修改。
5. 输出文件位于“output”标识的目录，具体文件名可以根据你的建议修改。
6. 项目在开发过程中每个版本进行备份，位于“bak”目录，版本号的规则由你来定。
7. 可以在项目文件夹内新建.enve环境文件夹，并建立对应的python子环境，把必要的组间和依赖都放在这里；你来做决定是否需要.enve环境文件夹，如果采用这条路，请使用“include-system-site-packages = true”参数，且一定要在Readme文档的开头注明运行方式（否则用户会被以全局python来执行），如果有多种运行方式，也需要写明不同命令及对应的功能。
8. 不用移动或修改工作目录下的“！ReadMe.txt”，也不需要参考这个文件的内容，这是我给自己的开发和使用提示，是我自己使用的文件。


开发时为当前项目生成一个 AGENTS.md 文件，方便未来作为AGENTS使用或者继续开发，在尽量不影响ai阅读的前提下，推荐使用中文作为解释，文件内容包含：
- Project overview, - Folder structure explanation, - Build and run commands, - Testing workflow, - Coding conventions, - Important files and entry points, - Any environment or dependency setup, - Make it concise but practical for an AI coding agent.



==================================

目标：
1. 
2. 
3. 

要求：
1. 
2. 
3. 

请直接生成完整项目结构


==================================

这是一个客户需求：

1. 输入：
2. 输出：
3. 附加：

请：
- 自行设计最优的项目结构。
- 写全部代码并调试，直到跑通。
- 写README、AGENTS.md、requirements.txt。


==================================



==================================



==================================



==================================


