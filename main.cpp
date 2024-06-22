#include <iostream>
#include <fstream>
using namespace std;

int main() {
    //声明一个文件输出类
    ifstream target;
    // ifstream的open()方法用于打开文件
    target.open("Fields.bas");
    if (!target)
        cout << "cannot open target files :( " << endl;
    else
        cout << "file can be read :)" << endl;
        target.close();
    return 0;
}
