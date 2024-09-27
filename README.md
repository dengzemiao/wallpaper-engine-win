# wallpaper-engine-win

- 简介

  Python 3.11.9

  windows 指定壁纸文件夹，定时切换桌面壁纸源码。

  当前打包 exe 后配置缓存文件会直接创建在当前软件目录，所以可以创建一个文件夹存放改 exe 软件，在发送一个快捷方式到桌面进行使用，这样可以避免缓存配置文件影响桌面美观。

- 依赖的插件

  - 使用 pystray 库来管理托盘图标

    ```sh
    $ pip install pystray
    ```

  - 每次双击打开程序时，如果已有实例在运行，退出

    ```sh
    $ pip install pywin32
    ```

- 打包

  - 打包插件安装

    ```sh
    $ pip install pyinstaller
    ```

  - 配置好环境变量

    - 右键点击“此电脑”或“我的电脑”，选择“属性”。
    - 点击“高级系统设置”，然后点击“环境变量”。
    - 在“系统变量”或“用户变量”中找到 Path，并添加上述路径。
    - 下面路径可以按下面的文件夹层级自行查找，需要包含 pyinstaller 执行文件的路径即可。
    - `C:\Users\Administrator\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\LocalCache\local-packages\Python311\Scripts`

  - 打包指令

    - 确保 `app_package_name` 与打包的 `--name "WallpaperSwitcher"` 保持一致，这样才能支持开机启动切换。

    ```sh
    $ pyinstaller --onefile --windowed --icon=icon.ico --add-data "icon.ico;." --name "WallpaperSwitcher" main.py
    ```

- 【辅助】确认开机启动是否设置成功：

  1、按 Win + R，输入 regedit 打开注册表编辑器。
  2、导航到以下路径，复制贴进路径覆盖即可：`HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run`
  3、在右侧窗格中，查找您添加的启动项名称。如果存在，说明设置成功。

- 待解决问题

  1、开机启动配置或文件无权限
  2、托盘菜单希望鼠标左击出现，而不是右击
