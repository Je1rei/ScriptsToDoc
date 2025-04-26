# Script Collector → Word  
Generate a single **.docx** that contains every C# script in your Unity project – ready to share with teammates, teachers, reviewers or clients.

![Unity](https://img.shields.io/badge/Unity-2020.3%2B-black?logo=unity)
![License](https://img.shields.io/github/license/your-name/ScriptCollectorToWord)
![OpenUPM](https://img.shields.io/badge/OpenUPM-coming--soon-blue)

---

## Features
| ✔ | What it does |
|---|--------------|
| **One-click export** | Collects all `*.cs` files (optionally only selected sub-folders) and saves them into a single Word document. |
| **Keeps file names** | Each script is preceded by its file name for quick navigation. |
| **Editor-only** | Lives entirely inside the Unity Editor – no code touches your build. |
| **Folder filters** | Include / exclude `Editor` folders and choose sub-folders manually. |
| **Dependency helper** | Shows missing-dependency hints for NuGetForUnity and `DocumentFormat.OpenXml`. |

---

## Quick Start

1. **Install NuGetForUnity**

   *Window → Package Manager → Add package from Git URL…*  
2. **Install Open XML SDK**  
*NuGet → Manage NuGet Packages* → search **DocumentFormat.OpenXml** → **Install**.

3. **Add ScriptCollectorToWord**  
*Option A* – copy the `Editor` folder into your project.  
*Option B* – add this repo as a Git dependency in **manifest.json**:
```jsonc
"com.yourname.script-collector-word": "https://github.com/your-name/script-collector-word.git#upm"
