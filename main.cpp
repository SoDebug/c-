#include <iostream>
#include <fstream>
#include <string>
# define user_input_filename false

using namespace std;

int main() {
    string file_name;
    string Header;

    if (user_input_filename){
        // 提示用户输入文件名
        cout << "请键入文件名" << endl;
        cin >> file_name;
    }
    else{
        file_name = "script";
    }
    file_name = file_name + ".bas";
    cout << file_name <<endl;

    // 创建并打开文件
    ofstream target(file_name.c_str());

    // 检查文件是否成功打开
    if (!target) {
        cout << "无法打开文件 " << file_name << " :(" << endl;
    } else {
        cout << "文件" << file_name << " 已创建 :)" << endl;
        target << "Attribute VB_Name = \"NewMacros\"" << endl;
        target << "Sub AutoDoc()" << endl;
        target << "\'" << endl;
        target << "\'  AutoDoc 宏" << endl;
        target << "\'" << endl;
        target << "    Selection.EndKey Unit:=wdStory" << endl;
        target << "    Selection.InsertBreak Type:=wdSectionBreakNextPage" << endl;
        target << "    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False" << endl;
        target << "    Selection.TypeText Text:=\"SET 项目代号 \"\"\"\"\"" << endl;
        target << "    Selection.MoveLeft Unit:=wdCharacter, Count:=1" << endl;
        target << "    Selection.TypeText Text:=\"1563\"" << endl;
        target << "    Selection.MoveRight Unit:=wdCharacter, Count:=3" << endl;
        target << "    Selection.TypeParagraph" << endl;
        target << "    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False" << endl;
        target << "    Selection.TypeText Text:=\"REF 项目代号\"" << endl;
        target << "    Selection.WholeStory" << endl;
        target << "    Selection.Fields.Update" << endl;
        target << "End Sub" << endl;
        target.close();  // 关闭文件
    }

    return 0;
}
