import win32com.client
import win32traceutil
import os, sys, imp
import pythoncom

class COM_test:
    _public_methods_ = ['where_am_i']
    _reg_progid_ = "PythonZ.COMTest"
    _reg_clsid_ = '{BED95BAF-13C4-40AC-92DF-543E452FE665}'    

    def main_is_frozen(self):                       #returns True when running the exe, 
        return (hasattr(sys, "frozen") or           #and False when running from a script
                hasattr(sys, "importers") or        
                imp.is_frozen("__main__"))          

    def get_main_dir(self):
        ''' get_main_dir() returns the directory name of the script (when running from the python script) 
            or the directory name of the exe (when running from the COM server)
            http://www.py2exe.org/index.cgi/HowToDetermineIfRunningFromExe  '''

        if self.main_is_frozen():
            print("main is frozen")
            return os.path.dirname(sys.executable)   #the absolute path of the executale binary of the Python Interpreter
        print("main is not frozen")
        return sys.path[0]            #the directory containing the script that was used to invoke the Python interpreter

    def where_am_i(self):
        dir = self.get_main_dir()
        print ("COM running from ", dir)
        return dir
        
if __name__=='__main__':
    print ("Registering COM Server")
    import win32com.server.register
    win32com.server.register.UseCommandLine(COM_test, debug=True)

    #COM = COM_test()
    #COM.where_am_i()