#include <iostream>
#include <fstream>
#include <string>
# define user_input_filename false

using namespace std;

int main() {
    string file_name;
    string Header;

    if (user_input_filename){
        // ��ʾ�û������ļ���
        cout << "������ļ���" << endl;
        cin >> file_name;
    }
    else{
        file_name = "script";
    }
    file_name = file_name + ".bas";
    cout << file_name <<endl;

    // ���������ļ�
    ofstream target(file_name.c_str());

    // ����ļ��Ƿ�ɹ���
    if (!target) {
        cout << "�޷����ļ� " << file_name << " :(" << endl;
    } else {
        cout << "�ļ�" << file_name << " �Ѵ��� :)" << endl;
        target << "Attribute VB_Name = \"NewMacros\"" << endl;
        target << "Sub AutoDoc()" << endl;
        target << "\'" << endl;
        target << "\'  AutoDoc ��" << endl;
        target << "\'" << endl;
        target << "    Selection.EndKey Unit:=wdStory" << endl;
        target << "    Selection.InsertBreak Type:=wdSectionBreakNextPage" << endl;
        target << "    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False" << endl;
        target << "    Selection.TypeText Text:=\"SET ��Ŀ���� \"\"\"\"\"" << endl;
        target << "    Selection.MoveLeft Unit:=wdCharacter, Count:=1" << endl;
        target << "    Selection.TypeText Text:=\"1563\"" << endl;
        target << "    Selection.MoveRight Unit:=wdCharacter, Count:=3" << endl;
        target << "    Selection.TypeParagraph" << endl;
        target << "    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False" << endl;
        target << "    Selection.TypeText Text:=\"REF ��Ŀ����\"" << endl;
        target << "    Selection.WholeStory" << endl;
        target << "    Selection.Fields.Update" << endl;
        target << "End Sub" << endl;
        target.close();  // �ر��ļ�
    }

    return 0;
}
