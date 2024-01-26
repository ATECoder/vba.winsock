# Cloning

- [Source Code](#Source-Code)
  - [Repositories](#Repositories)
  - [Global Configuration Files](#Global-Configuration-Files)
  - [Packages](#Packages)

<a name="Source-Code"></a>
## Source Code
Clone the repository along with its requisite repositories to their respective relative path.

### Repositories
The repositories listed in [external repositories] are required:

[Core repository]
[MVVM repository] (pending)

```
git clone https://github.com/ATECoder/vba.core.git
git clone https://github.com/ATECoder/vba.mvvm.git
git clone https://github.com/ATECoder/vba.winsock.git
```

Clone the repositories into the following folders (parents of the .git folder):
```
%vba%\core\core
%vba%\iot\winsock
```
where %vba% is the root folder of the VBA libraries, e.g., %my%\lib\vba, and %my%, e.g., c:\my is the overall root folder.

[external repositories]: ExternalReposCommits.csv
[Core repository]: https://github.com/ATECoder/vba.core.git
[MVVM repository]: https://github.com/ATECoder/vba.mvvm.git
