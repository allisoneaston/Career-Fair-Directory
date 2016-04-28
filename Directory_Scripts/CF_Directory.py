import os
import cgi
import csv
from django.conf import settings
from django.core.validators import validate_email
from django.core.validators import URLValidator
from django.core.exceptions import ValidationError


def get_degrees(degree_string):
    output=[]
    if degree_string.lower().find('bachelor')>-1:
        output.append('Bachelor\'s')
    if degree_string.lower().find('master')>-1:
        output.append('Master\'s')
    if degree_string.lower().find('doctor')>-1:
        output.append('Ph.D.')
    return output
#Does not support duplicate names yet
def sort_by_last_name(names):
    last_names ={}
    sorted_names = []
    for name in names:
        parts=name.split()
        if parts[len(parts)-1] in last_names:
            last_names[parts[len(parts)-1]].append(name)
        else:
            last_names[parts[len(parts)-1]]=[name]
    for last in sorted(last_names.keys()):
            for full_name in sorted(last_names[last]):
                sorted_names.append(full_name)
    return sorted_names
        

class ChairData:
    Committees={}
    Directors ={}
    Director_signatures = {}
    def __init__(self, chair_file,director_file):
        Committees = {}
        chair_file_w_path = os.path.join("..","Directory_Data",chair_file)
        director_file_w_path = os.path.join("..","Directory_Data",director_file)
        CommitteeReader = csv.reader(open(chair_file_w_path,'r'), delimiter=',')
        DirectorReader = csv.reader(open(director_file_w_path,'r'),delimiter=',')
        for row in CommitteeReader:
            Committees[row[0]]=[]
            for index in range(1,len(row)):
                if row[index]=="":
                    continue
                Committees[row[0]].append(row[index])
        self.Committees=Committees
        Directors = {}
        DirectorReader.next()
        for row in DirectorReader:
            Directors[row[0]]=row[1]
            self.Director_signatures[row[0]]=row[2]
        self.Directors = Directors
    
    def get_Chair_String(self):
        committee_string = r'''\begin{multicols}{2}
    '''
        for comm in sorted(self.Committees.keys()):
            committee_string+=r'''\begin{minipage}{\columnwidth}
    {\bf %(comm_name)s}\\
    '''%{"comm_name":comm}
            for chair in sort_by_last_name(self.Committees[comm]):
                committee_string+=r'''%(chair_name)s\\
    '''%{"chair_name":chair}
            committee_string+="\n"
            committee_string+=r'''\end{minipage}
    '''
        committee_string+=r'''\end{multicols}'''
        return committee_string
    
    def get_Director_String(self):
        director_string = r'''\begin{multicols}{2}
    '''
        for director in sort_by_last_name(self.Directors.keys()):
            file_name = self.Director_signatures[director]
            director_string+=(r'''\begin{minipage}{\columnwidth}
            \includegraphics[height=0.6in, width=2in,keepaspectratio]{'''+file_name+r'''}\\
            '''+director+r'''\\
            '''+self.Directors[director]+r''' Career Fair Director\\
            \end{minipage}'''+"\n")
        director_string+=r'''\end{multicols}'''
        return director_string

class MajorData:
    Major_Cipher = {"Aerospace Engineering":"AERO",
                "Aerospace Science":"AERO",
                "Automotive Engineering":"AUTO",
                'Atmospheric, Oceanic, and Space Sciences':"AOSS",
                "Biomedical Engineering":"BME",
                "Chemical Engineering":"CHE",
                "Plastics Engineering":"CHE",
                "Civil and Environmental Engineering":"CEE",
                "Civil Engineering":"CEE",
                "Construction Engineering and Management":'CEE',
                "Computer Engineering":"CE",
                "Space Engineering": "CLASP",
                "Space and Planetary Physics":"CLASP",
                "Climate and Space Sciences and Engineering":"CLASP",
                "Earth Systems Science and Engineering":"CLASP",
                'Applied Climate':"CLASP",
                "Applied Remote Sensing and Geoinformation Systems":"CLASP",
                "Computer Science":"CS",
                "Computer Science Engineering":"CS",
                "Computer Science and Engineering":"CS",
                "Data Science": "DS",
                'CS':'CS',
                'CE':'CE',
                "Electrical Engineering":"EE", 
                "Environmental Engineering":"ENV",
                "Engineering Physics":"EP",                            
                "Engineering Math/Physics":"EP",                            #Double check
                "Financial Engineering":"FE",
                "Industrial Engineering":"IOE",                                #Double check
                "Industrial and Operations Engineering":"IOE",
                'Engineering (Interdisciplinary Degree Program)':'IP',
                "Integrative Systems and Design":"ISD",
                "Design Science":"ISD",
                "Integrative Systems & Design":"ISD",
                "Manufacturing Engineering":"MFE",
                "Global Automotive and Manufacturing Engineering":"ISD",
                "Manufacturing":"MFE",
                "Materials Science and Engineering":"MSE",
                "Materials Science Engineering":"MSE",
                "Macromolecular Science and Engineering": "MSE",
                "Materials Science":"MSE",
                "Material Science":"MSE",
                "Mechanical Engineering":"ME",
                "Structural Engineering":"CIVIL",
                "Applied Mechanics":"ME",
                "Concurrent Marine Design":"NAME",
                "Naval Architecture and Marine Engineering":"NAME",
                "Nuclear Engineering and Radiological Sciences":"NERS",
                "Nuclear Engineering":"NERS",
                "Radiological Science":"NERS",
                "Nuclear Science":"NERS",
                "Pharmaceutical Engineering":"PHARM",
                "Robotics and Autonomous Vehicles":"ROB",
                "Robotics":"ROB",
                "Systems Engineering":"EE:S",        
                "Electrical Engineering Systems":"EE:S",  
                "Electrical Engineering-Systems":"EE:S",
                "Integrated Microsystems":"EE:S",
                "Integrated MicroSystems":"EE:S",
                'All Majors':'Any',
                'Bachelors':'Any',
                'Masters':'Any',
                'Doctoral':'Any',
                'Engineering (undeclared)':'Any',
                'Energy Systems Engineering':'ESE',
                'Applied Physics':'APhys',
                'Entrepreneurship':"ENT",
                'Systems Engineering and Design':'ISD',
                
                
    }
    Major_Acron_Cipher = {"AERO":"Aerospace Engineering",
                "AOSS":"Atmospheric, Oceanic, and Space Sciences",
                "AUTO":"Automotive Engineering",
                "BME":"Biomedical Engineering",
                "CHEME":"Chemical Engineering",
                "CIVIL":"Civil Engineering",
                "CE":"Computer Engineering",
                "CS":"Computer Science Engineering",
                "EE":"Electrical Engineering", 
                "EP":"Engineering Physics",                            
                "ENV":"Environmental Engineering",                        #Double check
                "FE":"Financial Engineering",
                "IOE":"Industrial and Operations Engineering",
                "IP":"Interdisciplinary Programs", 
                "MFE":"Manufacturing Engineering",
                "MSE":"Materials Science and Engineering",
                "ME":"Mechanical Engineering",
                "NAME":"Naval Architecture and Marine Engineering",
                "NERS":"Nuclear Engineering and Radiological Sciences",
                "PHARM":"Pharmaceutical Engineering",
                "EE Systems":"Electrical Engineering Systems",                                #Double check
                "Mechan":"Mechanical Engineering",
                "Mechanical":"Mechanical Engineering",
                "Materials Sci":"Materials Science and Engineering",
                "Naval Arc":"Naval Architecture and Marine Engineering",
    }
    def major_word2acron(self,major_list):
        major_acron = []
        for major_name in major_list:
            if not major_name in ['Bachelors','Masters','Doctoral']:
                major = major_name.replace('Bachelors ','').replace('Masters ','').replace('Doctoral ','').replace('PhD ','').strip()
            else:
                major = major_name
            if len(major)<1:
                continue
            if(major.find('Atmospheric')>-1):
                major = "Atmospheric, Oceanic, and Space Sciences"
            elif(major.find('Oceanic')>-1 or major.find('Space Sciences')>-1):
                continue
            try:
                major_acron.append(self.Major_Cipher[major.strip()])
            except KeyError:
                print " ".join(["MalFormed Major: >", major, "< -skipping."])
                
                
        return set(major_acron)
    
    def major_acron2word(self,major_list):
        major_word = []
        for major in major_list:
            if len(major)<1:
                continue
            try:
                major_word.append(self.Major_Acron_Cipher[major.strip()])
            except KeyError:
                print " ".join(["MalFormed Major: >", major, "< -skipping."])
                
                
        return set(major_word)
        

    
class SpreadSheetProcess:
    Indices ={}
    def get_Index(self,word):
        return self.Indices[word]
    def set_Indices(self,Top_Row):
        self.Indices = {"Name":-1,
               "Website":-1,
               "Majors":-1,
               "Positions":-1,
               "Degree":-1,
               "Location":-1,
               "Online-Application":-1,
               "Hiring-Policy":-1,
               "Attendance-Date":-1,
               "Receptions":-1,
               "Profile":-1,
               "Sponsor?":-1,
               "Building":-1}
        for index in range(len(Top_Row)):
            to_search=Top_Row[index].lower()
            if(self.Indices['Majors']<0): #Should work 2013
                if(to_search.find('major')>-1):
                    self.Indices['Majors']=index
                    print 'Majors: '+ str(index)
                    
                    
            if(self.Indices['Degree']<0): #Should work 2013
                if(to_search.find('degree')>-1):
                    self.Indices['Degree']=index
                    print 'Degree: '+ str(index)
                    continue
            if(self.Indices['Sponsor?']<0): #Should work 2013
                if(to_search.find('sponsorship')>-1):
                    self.Indices['Sponsor?']=index
                    print 'Sponsor?: '+ str(index)
                    continue
            if(self.Indices['Building']<0):
                if(to_search.find('building')>-1):
                    self.Indices['Building']=index
                    print 'Building: '+ str(index)
                    continue
            if(self.Indices['Receptions']<0): #Should work 2013
                if(to_search.find('reception')>-1):
                    self.Indices['Receptions']=index
                    print 'Receptions: '+ str(index)
                    continue
            if(self.Indices['Attendance-Date']<0): #Should work 2013
                if(to_search.find('date')>-1 or to_search.find('day')>-1):
                    self.Indices['Attendance-Date']=index
                    print 'Date: '+ str(index)
                    continue
            if(self.Indices['Hiring-Policy']<0): #Should work 2013
                if(to_search.find('willing to sponsor')>-1 or to_search.find('policy')>-1 or (to_search.find('hiring')>-1 and (to_search.find('number')>-1 or to_search.find('description')>-1))  ):
                    self.Indices['Hiring-Policy']=index
                    print 'Hiring-Policy: '+ str(index)
                    continue
            if(self.Indices['Online-Application']<0): #Should work 2013
                if(to_search.find('application')>-1 and (to_search.find('online') or to_search.find('web'))):
                    self.Indices['Online-Application']=index
                    print 'Online-Application: '+ str(index)
                    continue
            if(self.Indices['Website']<0): #Should work 2013
                if(to_search.find('web')>-1 and (self.Indices['Online-Application']>-1 or to_search.find('application')<0)):
                    self.Indices['Website']=index
                    print 'Wobsite: '+ str(index)
                    continue
            if(self.Indices['Location']<0): #Should work 2013
                if(to_search.find('location')>-1 or to_search.find('region')>-1 ):
                    self.Indices['Location']=index
                    print 'Region: '+ str(index)
                    continue
            if(self.Indices['Profile']<0): #Should work 2013
                if((to_search.find('profile')>-1 or to_search.find('description')>-1)and (self.Indices['Hiring-Policy']>-1 or to_search.find('hiring')<0) ):
                    self.Indices['Profile']=index
                    print 'profile: '+ str(index)
                    continue
            if(self.Indices['Positions']<0):#Should work 2013
                if(to_search.find('position')>-1 or (to_search.find('type')>-1 and (to_search.find('job')>-1 or to_search.find('employ')>-1 or to_search.find('opportunity')>-1)) ):
                    self.Indices['Positions']=index
                    print 'positions: '+ str(index)
                    continue
            if(self.Indices['Name']<0): #Should work 2013
                if(to_search.find('name')>-1):
                    self.Indices['Name']=index
                    print 'name: '+ str(index)
                    continue
        if min(self.Indices.values())<0:
            print 'WARNING: Unset indices'
        if(len(self.Indices.values())!=len(set(self.Indices.values()))):
            print 'WARNING: Duplicate indices'
    
class StringCleaning:
    def __init__(self):
        try:
            settings.configure()
        except RuntimeError:
            pass
    def YesNo2Bool(self,yn):
        if(yn.lower().find('y')>-1):
            return True
        else:
            return False
    def is_url(self,word):
        val = URLValidator()
        try: 
            val(word)
            return True
        except (ValidationError,ValueError):
            try:
                val('http://'+word)
                return True
            except (ValidationError,ValueError):
                return False
            return False
        # if word.startswith('http'):
            # return True
        # if word.startswith('www.'):
            # return True
        # if word.count('@')>0:
            # return False
        # if word.endswith('.com') or word.endswith('.edu') or word.endswith('.org'):
            # return True
        # if word.endswith('.gov') or word.endswith('.html') or word.endswith('.cfm'):
            # return True
        # if word.endswith('.htm') or word.endswith('.ca') or word.endswith('.net'):
            # return True
        # return False
    def is_email(self,word):
        try: 
            validate_email(word)
            return True
        except (ValidationError,ValueError):
            return False
    

    def clean_string_latex(self,input):
        #output=unicode(input,encoding="utf-8",errors="replace")
        output=input.replace('%','\%')
        output=output.replace('$','\$')
        output=output.replace('&','\&')
        output=output.replace('_','\_')
        output=output.replace('#',r'''\#''')
        output_list=output.split()
        count=0
        while count<len(output_list):
            word=output_list[count];
            if self.is_url(word):
                output_list[count]=r'''\url{'''+word+r'''}'''
            if self.is_email(word):
                output_list[count]=r'''\href{mailto:'''+word+r'''}{\nolinkurl{'''+word+r'''}}'''
            count=count+1
        output=' '.join(output_list)
        return output
        
    def clean_string_web(self,input):
        temp1 = cgi.escape(input,quote=True)
        output=temp1.encode("ascii",'xmlcharrefreplace')
        #output = cgi.escape(input, quote=True).encode("ascii",'xmlcharrefreplace')
        return output
    def clean_string_js(self,input):
        output=self.clean_string_web(input).replace("'",r'''\'''')
        return output
        
