// JavaScript Document
function getlocal(){   
    //得到本地路径   
    temp = window.location+".txt";   
    temp = temp.substr(8);   
    temp=temp.split("/")   
    temp.pop()   
    str=temp.join("/")+"/"  
    return str   
}   
  
  
function writefile(name, str, always){   
    //写文件[name:文件名,str:字符,always:是否覆盖]   
    //用法:writefile("文件名", "新文件", true)   
    var fso, f;   
    if (always) {   
        fso = new ActiveXObject("Scripting.FileSystemObject");   
        f = fso.OpenTextFile(getlocal()+name, 2, true);   
        f.Write(str);   
        f.Close();   
        return true;   
    } else {   
        if (!isfile(name)){   
            fso = new ActiveXObject("Scripting.FileSystemObject");   
            f = fso.OpenTextFile(getlocal()+name, 2, true);   
            f.Write(str);   
            f.Close();   
            return true;   
        } else {   
            return false;   
        }   
    }   
}   
  
  
function readfile(name) {   
    //读取一个文件,如果存在返回字符,如不存在返回flash   
    if (isfile(name)) {   
        var fso, f;   
        fso = new ActiveXObject("Scripting.FileSystemObject");   
        f = fso.OpenTextFile(getlocal()+name, 1);   
        str = f.ReadAll();   
        f.Close();   
        return (str);   
    } else {   
        return false;   
    }   
}   
  
  
function createfolder(name){   
    //创建一个文件夹   
    if (!isfolder(name)) {   
        var fso, f;   
        fso = new ActiveXObject("Scripting.FileSystemObject");   
        fso.CreateFolder(getlocal()+name);   
        return true;   
    } else {   
        return false;   
    }   
}   
  
  
function deletefolder(name) {   
    //删除一个文件夹   
    if (isfolder(name)) {   
        var fso, f;   
        fso = new ActiveXObject("Scripting.FileSystemObject");   
        fso.DeleteFolder(getlocal()+name);   
        return true  
    } else {   
        return false;   
    }   
}   
  
  
function isfile(name){   
    //文件是否存在   
    var fso, exis;   
    fso = new ActiveXObject("Scripting.FileSystemObject");   
    exis=fso.FileExists(getlocal()+name);    
    return exis   
}   
  
  
function isfolder(name){   
    //文件夹是否存在   
    var fso, exis;   
    fso = new ActiveXObject("Scripting.FileSystemObject");   
    exis=fso.FolderExists(getlocal()+name);    
    return exis   
}   
  
  
function viewfolder(name) {   
    //查看一个文件夹并返回所有文件的名   
    if (isfolder(name)) {   
        var fso, f;   
        fso = new ActiveXObject("Scripting.FileSystemObject");   
        f = fso.GetFolder(getlocal()+name).files;   
        str = "";   
        for (var objEnum = new Enumerator(f); !objEnum.atEnd(); objEnum.moveNext())   
        {   
            temp = String(objEnum.item()).split("\\").pop();   
            str += temp+"|";   
               
        }   
        return str   
    } else {   
        return false;   
    }   
}  