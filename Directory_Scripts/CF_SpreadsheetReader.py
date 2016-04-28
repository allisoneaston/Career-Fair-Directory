import argparse
import codecs
import csv
from datetime import datetime
from math import ceil,pow,sqrt
import os
import subprocess
import sys
from CF_Directory import get_degrees
import CF_Directory

major_data=CF_Directory.MajorData()
spreadsheet=CF_Directory.SpreadSheetProcess()
string_cleaner = CF_Directory.StringCleaning()
personnel = None
dates=['Monday September 28','Tuesday September 29']
date_strings = ['Monday, September 28','Tuesday, September 29']
#Make sure all text files are saved using utf-8 encoding

def get_hiring_policy(policy_string):
    if policy_string.lower().find('yes')>-1:
        return 1
    if policy_string.lower().find('no')>-1:
        return 3
    return 2
def get_date(date_string):
    if date_string.find('Day 1')>-1:
        return dates[0]
    return dates[1]
def unicode_csv_reader(windows_data, dialect=csv.excel, **kwargs):
    csv_reader = csv.reader(windows_data, dialect=dialect, **kwargs)
    for row in csv_reader:
        yield [unicode(cell, 'latin-1') for cell in row]

def get_cf_number(year):
    numeral = year-1984
    if numeral%100 in range(11,13):
        return r'''%d\textsuperscript{th}'''%(numeral)
    if numeral%10 ==1:
        return r'''%dtextsuperscript{st}'''%(numeral)
    elif numeral%10==2:
        return r'''%dtextsuperscript{nd}'''%(numeral)
    elif numeral%10==3:
        return r'''%dtextsuperscript{rd}'''%(numeral)
    else:
        return r'''%d\textsuperscript{th}'''%(numeral)

def company_blurb_short(Company):
    if string_cleaner.is_url(Company["ONLINE-APPLICATION"]):
        web = Company["ONLINE-APPLICATION"]
    else:
        web = Company["WEBSITE"]
    return r'''\begin{tabularx}{.95\columnwidth}{Xr}
                 {\Large\bf %(name)s} & {\Large\bf %(building)s}\\
    \multicolumn{2}{p{.95\columnwidth}}{\url{%(web)s}}\\
    \multicolumn{2}{p{.95\columnwidth}}{\emph{Majors:} %(majors)s}\\
    \multicolumn{2}{p{.95\columnwidth}}{\emph{Positions:} %(positions)s}\\
    \multicolumn{2}{p{.95\columnwidth}}{\emph{Degrees:} %(degrees)s}\\
    \multicolumn{2}{p{.95\columnwidth}}{\emph{Hiring Policy:} %(hire-policy)d}\\
    \end{tabularx}
    '''%{"name":Company["NAME"],"web":web.replace('\\','/').replace('#','\\#'),"majors":", ".join(Company["MAJOR_ACRON"]),"positions":", ".join(Company["POSITIONS"]),"degrees":",".join(Company["DEGREES"]),"hire-policy":Company["HIRING-POLICY"],"building":Company["BUILDING"]}

def company_blurb_long(Company):
    if string_cleaner.is_url(Company["ONLINE-APPLICATION"]):
        web = Company["ONLINE-APPLICATION"]
    else:
        web = Company["WEBSITE"]
    return r'''{\Large\bf %(name)s \hfill %(building)s}\\
    \url{%(web)s}\\
    %(description)s\\
    \emph{Majors:} %(majors)s\\
    \emph{Positions:} %(positions)s\\
    \emph{Degrees:} %(degrees)s\\
    \emph{Hiring Policy:} %(hire-policy)d\\'''%{"name":Company["NAME"],"web":web.replace('\\','/').replace('#','\\#'),"majors":", ".join(Company["MAJOR_ACRON"]),"positions":",".join(Company["POSITIONS"]),"degrees":",".join(Company["DEGREES"]),"hire-policy":Company["HIRING-POLICY"],'description':Company["PROFILE"],"building":Company["BUILDING"]}

def make_binder_covers(Monday,Tuesday):
    gen_binder_path="BinderCovers.tex"
    gen_binder_covers=open(gen_binder_path,'w')
    gen_binder_covers.write(r'''\documentclass{article}
    \input{../Directory_Tex_Files/AuxilliaryStuff.tex}
    \setupdirectory
    %\usepackage[margins=10pt]{geometry}
    \newgeometry{margin=0.75in}
    \input{../Directory_Tex_files/BinderCover.tex}
    \begin{document}''')
    for company in Monday:
        name = company["NAME"]
        if len(name) >50:
            name = r'''{\fontsize{30}{36}\selectfont '''+name+'}'
        gen_binder_covers.write(r'''\companybindercover{%(name)s}
    '''%{'name':name})
    for company in Tuesday:
        name = company["NAME"]
        if len(name) >50:
            name = r'''{\fontsize{30}{36}\selectfont '''+name+'}'
        gen_binder_covers.write(r'''\companybindercover{%(name)s}
    '''%{'name':name})
    gen_binder_covers.write(r'''\end{document}''')
    gen_binder_covers.close()
    
def make_html_page(Monday,Tuesday):    
    gen_html_path="CareerFairCompanySheet.html"
    gen_html=open(gen_html_path,'w')
    gen_html.write(r'''<!doctype html>
    <html lang="en">

    <!--Meta information-->
    <head>
    <meta charset=utf-8>
    <title>SWE/TBP Career Fair: Companies Attending</title>
    <link rel="stylesheet" href="static/css/bootstrap2.css" type="text/css">
    <link rel="stylesheet" href="static/css/cf.css" type="text/css">
    <meta name=description value="The SWE/TBP Career Fair held annually at the University of Michigan College of Engineering">
    <meta name=keywords value="job, career, fair, swe, engineering, tbp, michigan">

    <!--[if lt IE 9]>
    <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
    <![endif]-->

    <script type="text/javascript">
    var pagename="Inventory";
    </script>
    <!--<script type="text/javascript">

      var _gaq = _gaq || [];
      _gaq.push(['_setAccount', 'UA-26073977-3']);
      _gaq.push(['_trackPageview']);

      (function() {
        var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
        ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
        var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
      })();

    </script>-->
    </head>

    <!--Main Content-->
    <body>

    <article>
    <h1>Companies Attending</h1>
    <p>Click on column headers to sort the table by that column. Type in boxes to filter. For company descriptions, hover over the company name.</p>
    <h2>List of Majors</h2>
  <table>
    <tr>
      <td><strong>AERO:</strong> Aerospace Engineering</td>
      <td><strong>AOSS:</strong> Atmospheric, Oceanic, Space Sciences</td>
      <td><strong>APhys:</strong> Applied Physics</td>
      <td><strong>AUTO:</strong> Automotive Engineering</td>
    </tr>
    <tr>
      <td><strong>BME:</strong> Biomedical Engineering</td>
      <td><strong>CE:</strong> Computer Engineering</td>
      <td><strong>CHEME:</strong> Chemical Engineering</td>
      <td><strong>CIVIL:</strong> Civil Engineering</td>
    </tr>
    <tr>
      
      <td><strong>CS:</strong> Computer Science</td>
      <td><strong>EE:</strong> Electrical Engineering</td>
      <td><strong>EE:S:</strong> Electrical Engineering Systems</td>
      <td><strong>ENT:</strong> Entrepreneurship</td>
    </tr>
    <tr>
      <td><strong>ENV:</strong> Environmental Engineering</td>
      <td><strong>EP:</strong> Engineering Physics</td>
      <td><strong>ESE:</strong> Energy Systems Engineering</td>
      <td><strong>FE:</strong> Financial Engineering</td>
    </tr>
    <tr>
      <td><strong>IOE:</strong> Industrial and Operations Engineering</td>
      <td><strong>IP:</strong> Interdisciplinary Programs</td>
      <td><strong>ISD:</strong> Integrative Systems+Design</td>
      <td><strong>MECHE:</strong> Mechanical Engineering</td>
    </tr>
    <tr>
      <td><strong>MFE:</strong> Manufacturing Engineering</td>
      <td><strong>MSE:</strong> Materials Science and Engineering</td>
      <td><strong>NAME:</strong> Naval Architecture and Marine Engineering</td>
      <td><strong>NERS:</strong> Nuclear Engineering and Radiological Sciences</td>
    </tr>
    <tr>
      <td><strong>PHARM:</strong> Pharmaceutical Engineering</td>
      <td><strong>ROB:</strong> Robotics</td>
    </tr>
  </table>

    <h2>Hiring Policy:</h2>
    <p>Is the company willing to sponsor selected candidates for work authorization?</p>
    <ol>
    <li>Yes</li>
    <li>On Occassion</li>
    <li>No</li>
    </ol>
    <div class="table-holder">
    <table class=" table table-striped table-autosort:0" id="header-fixed">
    <thead style="background: white;"><tr>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by company name">Company</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by majors Hhring">Majors Hiring</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by positions available">Positions Available</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by degree required">Degree Required</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by location">Job Location</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by attendance date">Date Attending</th>
    <th class="table-sortable:numeric table-filterable" title="Click to sort by hiring policy">Hiring Policy</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by expected receptions attendance">Attending Receptions?</th>
    <th class="table-sortable:ignorecase table-filterable" title="Click to sort by application webpage">Online Application Web Page</th>
    </tr>
    <tr>
    <th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
    <th><input name="filter" size="8" onkeyup="Table.filter(this,this)"></th>
    <th><select onchange="Table.filter(this,this)">
        <option value="function(){return true;}">Any</option>
        <option value="full">Full-Time</option>
        <option value="intern">Summer Internship</option>
        <option value="co.*op">Co-Op</option>
        </select></th>
    <th><select onchange="Table.filter(this,this)">
        <option value="function(){return true;}">Any</option>
        <option value="ph">Ph.D.</option>
        <option value="master">Master's</option>
        <option value="bachelor">Bachelor's</option>
        </select></th>
    <th><select onchange="Table.filter(this,this)">
        <option value="function(){return true;}">Any</option>
        <option value="west.*coast">West Coast</option>
        <option value="southwest">Southwest</option>
        <option value="northwest">Northwest</option>
        <option value="south.*east">Southeast</option>
        <option value="northeast">Northeast</option>
        <option value="midwest">Midwest</option>
        </select></th>
    <th><select onchange="Table.filter(this,this)">
        <option value="function(){return true;}">All Days</option>
        <option value="mon">Monday</option>
        <option value="tues">Tuesday</option>
        </select></th>
    <th><select onchange="Table.filter(this,this)">
        <option value="function(){return true;}">All policies</option>
        <option value="function(val){return parseFloat(val)<=1;}">Yes</option>
        <option value="function(val){return parseFloat(val)<=2;}">On occassion</option>
        <option value="function(val){return parseFloat(val)<=3;}">No</option>
        </select></th>
    <th id="centered-cell"><select onchange="Table.filter(this,this)">
        <option value="function(){return true;}">Either</option>
        <option value="yes">Yes</option>
        <option value="no">No</option>
        </select></th>
    <th></th>
    </tr>
    </thead><tbody>''')

    for company in (Monday+Tuesday):
        attending_receptions = "Yes" if company["RECEPTIONS"] else "No"
        company_website = company["WEBSITE"] if company["WEBSITE"].find('http:')>=0 else "http://"+company["WEBSITE"]
        words_list = company["ONLINE-APPLICATION"].replace('\n',' ').split(' ')
        for count in range(len(words_list)):
            words_list[count]= string_cleaner.clean_string_web(words_list[count])
            if string_cleaner.is_url(words_list[count]):
                if not words_list[count].startswith('http'):
                    words_list[count]='http://'+words_list[count]
                words_list[count]=r'''<a href="'''+words_list[count]+r'''">Online Application</a>'''
            elif string_cleaner.is_email(words_list[count]):
                words_list[count]=r'''<a href="mailto:'''+words_list[count]+r'''">'''+words_list[count]+r'''</a>'''

        online_application_string = ' '+(' '.join(words_list))
        description_string = string_cleaner.clean_string_js(company["WEB_PROFILE"])

        #print description_string
        #print len(description_string)
        #if len(description_string)>925:
        #    print 'Char'+description_string[926]
        #print company["WEB_NAME"]
        gen_html.write(r'''<td><a href="%(website)s" rel="tooltip" data-trigger="hover" data-toggle="popover" data-placement="right" title data-original-title="%(name)s" data-content="%(description)s">%(name)s</a></td>
                        <td>%(majors)s</td>
                        <td>%(positions)s</td>
                        <td>%(degrees)s</td>
                        <td>%(locations)s</td>
                        <td>%(date)s</td>
                        <td id="centered-cell">%(hiring)s</td>
                        <td id="centered-cell">%(receptions)s</td>
                        <td>%(application)s</td>
                        '''%{"name":string_cleaner.clean_string_web(company["WEB_NAME"]),
                            "website":company_website,
                            "positions":", ".join(company["POSITIONS"]),
                            "degrees":", ".join(company["DEGREES"]),
                            "majors":", ".join(["MECHE" if major=="ME" else major for major in company["MAJOR_ACRON"]]).replace(r'''$\mu$''',"&#956"),
                            "locations":", ".join(company["LOCATION"]),
                            "application":online_application_string,
                            "hiring":company["HIRING-POLICY"],
                            "date":company["ATTENDANCE-DATE"],
                            "receptions":attending_receptions,
                            "description":description_string})
        
        gen_html.write("</tr>")


    gen_html.write(r'''</tbody>
    </table>
    </div>

    </article>
    </body>
    <script src="static/js/sort_filter_table.js"></script>
    <script type="text/javascript" language="javascript" src="static/js/jquery-1.7.2.min.js"></script>
    <script type="text/javascript" language="javascript" src="static/js/sticky-header.js"></script>
    <script type="text/javascript" language="javascript" src="static/js/bootstrap2.min.js"></script>
    <script type="text/javascript">
    $(function(){
        $("[rel='tooltip']").popover()
    });
    </script>
    </html>''')

def make_directory(Monday,Tuesday,Monday_Majors,Tuesday_Majors,Sponsors,args):
    global personnel
    
    company_tex  = os.path.join("..","Directory_Tex_Files",args.output)
    
    administration_sheet=os.path.join("..","Directory_Data",args.administration)
    ECRC_sheet=os.path.join("..","Directory_Data",args.ecrc)
    facilities_sheet=os.path.join("..","Directory_Data",args.facilities)
    
    #CSV Readers
    AdministrationReader = unicode_csv_reader(open(administration_sheet,'r'), delimiter=',')
    ECRCreader = unicode_csv_reader(open(ECRC_sheet,'r'), delimiter=',')
    FacilitiesReader = unicode_csv_reader(open(facilities_sheet,'r'), delimiter=',')
    
    Directory_tex = codecs.open(company_tex,mode='w',encoding="utf-8")
    
    personnel = CF_Directory.ChairData(args.chairs,args.directors)
    header=r'''\documentclass[twoside]{article}
    \input{../Directory_Tex_Files/AuxilliaryStuff.tex}
    \setupdirectory
    \begin{document}
    \input{../Directory_Tex_Files/FrontPage.tex}
    '''
    tail = r'''\end{document}'''
    company_by_date_header=r'''\startcompanysection
    \fancyhead[C]{ {\fontspec{Bookman Old Style} \fontsize{40}{48}\selectfont \bf\color{white}Career Fair %(year_var)d}\\[12pt]
    {\fontspec{Monotype Corsiva}\fontsize{36}{41}\selectfont \color{white}\em %(date_var)s}\\
    }
    '''%{"year_var":datetime.now().year, "date_var":date_strings[0]}
    
    #Start and end different sections
    start_company= r'''\begin{center}\begin{multicols}{3}
    \begin{FlushLeft}
    '''
    start_sponsors=r'''\begin{center}\begin{multicols}{2}
    '''
    end_company = r'''\end{FlushLeft}
    \end{multicols}\end{center}
    '''
    end_sponsors=r'''\end{multicols}\end{center}
    '''
    start_majors = r'''\begin{center}\begin{multicols}{5}
    '''

    end_majors = r'''\end{multicols}\end{center}
    '''
    
    #Loads the Director Welcome Letter
    Director_Welcome_Letter=codecs.open(os.path.join("..","Directory_Text",args.directorletter),mode='r',encoding='utf-8-sig').read()%{"cf_number":get_cf_number(datetime.now().year)}
    
    #Starts Writing to the .tex file
    print "Writing to: "+args.output
    Directory_tex.write(header)
    Directory_tex.write(r'''\startforewardsection
    {\fontspec{Bookman Old Style} \fontsize{16}{19}\selectfont \bf Welcome}\\~\\''')
    
    #Director Letter Page
    Directory_tex.write(Director_Welcome_Letter.replace('\n',r'''~\\

    '''))
    Directory_tex.write(personnel.get_Director_String())
    Directory_tex.write(personnel.get_Chair_String())
    Directory_tex.write(r'''\newpage''')
    
    #Sponsor Letters
    for count in range(len(args.sponsornames)):
        DirectorySponsor = args.sponsornames[count]

        letter_text = codecs.open(os.path.join("..","Directory_Text",args.sponsorletters[count]),encoding='utf-8-sig').read()     
        num_breaks=letter_text.count('\n')
        num_lines=len(letter_text)/90+num_breaks
        if num_lines<52:
            fontsize=12
            line_skip=-.25
        else:
            fontsize=11*pow(52.0/num_lines,.95)
            line_skip=-.8
        SponsorLetter = string_cleaner.clean_string_latex(letter_text.replace('\n',r'''\\[%fex]~ '''%(line_skip)))
        SponsorLogo = args.sponsorlogos[count]
        Directory_tex.write(r'''{\fontspec{Bookman Old Style} \fontsize{20}{24}\selectfont \bf Message from Corporate Sponsor\\''')
        Directory_tex.write(DirectorySponsor+r'''}\\~\\''')
        num_lines = len(SponsorLetter)/90+num_breaks

        
        Directory_tex.write(r'''{\fontsize{%f}{%f}\selectfont '''%(fontsize,1*fontsize)+SponsorLetter+'}')
        Directory_tex.write(r'''\begin{center}\begin{adjustbox}{max size={.5\textwidth}{2in}}\includegraphics[height=2in]{'''+SponsorLogo+r'''}\end{adjustbox}\end{center}''')
        Directory_tex.write(r'''\newpage''')
    
    #Map, Major List, Locations and Hiring Policy, --doesn't change much
    Directory_tex.write(r'''\input{../Directory_Tex_Files/MapAndShit.tex}''')

    #Monday Companies--Sponsors
    Directory_tex.write(company_by_date_header)
    Directory_tex.write(r'''\startsponsorsection''')
    Directory_tex.write(r'''\fancyhead[C]{ {\fontspec{Bookman Old Style} \fontsize{40}{48}\selectfont \bf\color{white}Career Fair %(year_var)d}\\[12pt]
    {\fontspec{Monotype Corsiva}\fontsize{36}{41}\selectfont \color{white}\em %(date_var)s}\\
    \begin{tikzpicture}[remember picture,overlay]
    \node[yshift=-\headheight-10pt] at (current page.north)
    {\fontsize{16}{18}\selectfont \bf Sponsors};
    \end{tikzpicture}
    }
    '''%{"year_var":datetime.now().year, "date_var":date_strings[0]})
    Directory_tex.write(start_sponsors)
    for company in Monday:
        if(company["SPONSOR"]):
            Directory_tex.write('\\begin{minipage}{.95\columnwidth}')
            Directory_tex.write(company_blurb_long(company))
            Directory_tex.write('\n\\end{minipage}\n ')
            
    Directory_tex.write(end_sponsors)
    
    #Monday Companies--non sponsors
    Directory_tex.write(company_by_date_header)
    Directory_tex.write(start_company)
    for company in Monday:
        Directory_tex.write('\\begin{minipage}{\columnwidth}')
        Directory_tex.write(company_blurb_short(company))
        Directory_tex.write('\n\\end{minipage}\n \n')
    Directory_tex.write(end_company)
    Directory_tex.write(r'''\newpage''')
    
    #Monday Companies -- by major
    Directory_tex.write(r'''\newgeometry{left=0pt,right=15pt,bottom=8pt,includehead,headheight=1.6in,top=0in,twoside=false}
    \fancyhead[C]{ {\fontspec{Bookman Old Style} \fontsize{40}{48}\selectfont \bf\color{white}Career Fair %(year_var)d}\\[12pt]
    {\fontspec{Monotype Corsiva}\fontsize{36}{41}\selectfont \color{white}\em %(date_var)s}\\
    \begin{tikzpicture}[remember picture,overlay]
    \node[yshift=-\headheight-10pt] at (current page.north)
    {\fontsize{16}{18}\selectfont \bf Companies by Major};
    \end{tikzpicture}
    }
    '''%{"year_var":datetime.now().year, "date_var":date_strings[0]})
    Directory_tex.write(start_majors)
    for major in sorted(Monday_Majors.keys()):
        Directory_tex.write(r'''{\fontsize{14}{16}\selectfont \bf %(major)s}\\
        \vspace{-1em}
        ~\hrulefill~
        \vspace{-.9em}
        '''%{'major':major})
        Directory_tex.write(r'''\begin{FlushLeft}
        \begin{compactitem}
        ''')
        for company in Monday_Majors[major]:
            Directory_tex.write(r'''\item ''')
            Directory_tex.write(company+"\n")
        Directory_tex.write(r'''\end{compactitem}
        \end{FlushLeft}
        \vspace{1em}
        ''')
    Directory_tex.write(end_majors)
    Directory_tex.write(r'''\newpage
    ''')
    
    ## Tuesday Companies--Sponsors First
    Directory_tex.write(company_by_date_header)
    Directory_tex.write(r'''\startsponsorsection''')
    Directory_tex.write(r'''\fancyhead[C]{ {\fontspec{Bookman Old Style} \fontsize{40}{48}\selectfont \bf\color{white}Career Fair %(year_var)d}\\[12pt]
    {\fontspec{Monotype Corsiva}\fontsize{36}{41}\selectfont \color{white}\em %(date_var)s}\\
    \begin{tikzpicture}[remember picture,overlay]
    \node[yshift=-\headheight-10pt] at (current page.north)
    {\fontsize{16}{18}\selectfont \bf Sponsors};
    \end{tikzpicture}
    }
    '''%{"year_var":datetime.now().year, "date_var":date_strings[1]})
    Directory_tex.write(start_sponsors)
    for company in Tuesday:
        if(company["SPONSOR"]):
            Directory_tex.write('\\begin{minipage}{.95\columnwidth}')
            Directory_tex.write(company_blurb_long(company))
            Directory_tex.write('\n\\end{minipage}\n ')
            
    Directory_tex.write(end_sponsors)
    
    #Tuesday Companies--non sponsors
    Directory_tex.write(company_by_date_header)
    Directory_tex.write(r'''\fancyhead[C]{ {\fontspec{Bookman Old Style} \fontsize{40}{48}\selectfont \bf\color{white}Career Fair %(year_var)d}\\[12pt]
    {\fontspec{Monotype Corsiva}\fontsize{36}{41}\selectfont \color{white}\em %(date_var)s}\\
    }
    '''%{"year_var":datetime.now().year, "date_var":date_strings[1]})
    Directory_tex.write(start_company)
    for company in Tuesday:
        Directory_tex.write('\\begin{minipage}{.9\columnwidth}')
        Directory_tex.write(company_blurb_short(company))
        Directory_tex.write('\n\\end{minipage}\n \n')
    Directory_tex.write(end_company)
    Directory_tex.write(r'''\newpage''')

    #Tuesday Companies--by major
    Directory_tex.write(r'''\newgeometry{left=0pt,right=15pt,bottom=8pt,includehead,headheight=1.6in,top=0in,twoside=false}
    \fancyhead[C]{ {\fontspec{Bookman Old Style} \fontsize{40}{48}\selectfont \bf\color{white}Career Fair %(year_var)d}\\[12pt]
    {\fontspec{Monotype Corsiva}\fontsize{36}{41}\selectfont \color{white}\em %(date_var)s}\\
    \begin{tikzpicture}[remember picture,overlay]
    \node[yshift=-\headheight-10pt] at (current page.north)
    {\fontsize{16}{18}\selectfont \bf Companies by Major};
    \end{tikzpicture}
    }
    '''%{"year_var":datetime.now().year, "date_var":date_strings[1]})
    Directory_tex.write(start_majors)
    for major in sorted(Tuesday_Majors.keys()):
        Directory_tex.write(r'''{\fontsize{14}{16}\selectfont \bf %(major)s}\\
        \vspace{-1em}
        ~\hrulefill~
        \vspace{-.9em}
        '''%{'major':major})
        Directory_tex.write(r'''\begin{FlushLeft}
        \begin{compactitem}
        ''')
        for company in Tuesday_Majors[major]:
            Directory_tex.write(r'''\item ''')
            Directory_tex.write(company+"\n")
        Directory_tex.write(r'''\end{compactitem}
        \end{FlushLeft}
        \vspace{1em}
        ''')
    Directory_tex.write(end_majors)

    #Write Thank you page
    Directory_tex.write(r'''\startforewardsection
    {\fontspec{Bookman Old Style} \fontsize{16}{19}\selectfont \bf Thank You!}\\''')
    Directory_tex.write(r'''The directors of the %(year)s Society of Women Engineers 
    and Tau Beta Pi Career Fair would like to thank the following individuals and 
    companies for their contribution and continued support in organizing this 
    successful event:\\
    '''%{"year":datetime.now().year})
    
    #Thank administration
    Directory_tex.write(r'''
    {\fontspec{Bookman Old Style} \fontsize{14}{17}\selectfont \bf College of Engineering Administration and Staff}\\''')
    AdministrationReader.next()
    Directory_tex.write(r'''\begin{tabular}{p{2in}l}
    ''')
    for row in AdministrationReader:
        Directory_tex.write(row[0]+'\t&\t'+row[1]+r'''\\
        ''')
    Directory_tex.write(r'''\end{tabular}\\[1em]
    ''')
    #Thank ECRC
    Directory_tex.write(r'''
    {\fontspec{Bookman Old Style} \fontsize{14}{17}\selectfont \bf College of Engineering Career Resource Center}\\''')
    ECRCreader.next()
    Directory_tex.write(r'''\begin{tabular}{p{2in}l}
    ''')
    for row in ECRCreader:
        Directory_tex.write(row[0]+'\t&\t'+row[1]+r'''\\
        ''')
    Directory_tex.write(r'''\end{tabular}\\[1em]
    ''')
    #Thank Facilities
    Directory_tex.write(r'''
    {\fontspec{Bookman Old Style} \fontsize{14}{17}\selectfont \bf Facilities Coordination}\\''')
    FacilitiesReader.next()
    Directory_tex.write(r'''\begin{tabular}{p{2in}l}
    ''')
    for row in FacilitiesReader:
        Directory_tex.write(row[0]+'\t&\t'+row[1]+r'''\\
        ''')
    Directory_tex.write(r'''\end{tabular}\\[1em]
    ''')
    #Thank Sponsors
    Directory_tex.write(r'''
    {\fontspec{Bookman Old Style} \fontsize{14}{17}\selectfont \bf Corporate Sponsors (continued on next page)}
    \vspace{-1em}\begin{multicols}{2}''')
    for company in Sponsors:
        Directory_tex.write(company+r'''\\
        ''')
    Directory_tex.write(r'''\end{multicols}
    ''')
    
    Directory_tex.write(r'''
    {\fontspec{Bookman Old Style} \fontsize{14}{17}\selectfont \bf Special thanks to all student volunteers who assited with the event!}''')
    
    Directory_tex.write(r'''\newpage
    ''')
    #Write Blank Notes Page
    #This is defined in the auxiliary latex file, 
    #since it needs to determine if it needs to add
    #one or two pages (depends on the number of pages used)
    Directory_tex.write(r'''\notepages
    ''')
    #Diamond Sponsors Page
    Directory_tex.write(r'''{\fontspec{Bookman Old Style} \fontsize{16}{19}\selectfont \bf Thank You to Our Diamond Corporate Sponsors!}\\
    \begin{center}\vspace{-1.3em}
    ''')
    for logo in args.diamondlogos:
        Directory_tex.write(r'''\begin{adjustbox}{max size={0.65\textwidth}{!}}\includegraphics[height='''+str(.9/len(args.diamondlogos))+r'''\textheight]{'''+logo+r'''}\end{adjustbox}\vspace{1em}\\
        ''')
    Directory_tex.write(r'''\end{center}
    \newpage
    ''')
    #Platinum/Gold Page
    Directory_tex.write(r'''{\fontspec{Bookman Old Style} \fontsize{16}{19}\selectfont \bf Thank You to Our Platinum Corporate Sponsors!}
    \vspace{-.75em}\begin{center}
    ''')
    max_height=0.35/ceil(len(args.platinumlogos)/3)
    for count in range(len(args.platinumlogos)):
        logo=args.platinumlogos[count]
        Directory_tex.write(r'''\begin{minipage}{0.45\textwidth}\begin{center}\vfill\begin{adjustbox}{max size={!}{'''+str(max_height)+r'''\textheight}}\includegraphics[width=\textwidth]{'''+logo+r'''}\end{adjustbox}\vfill\end{center}\end{minipage}
        ''')
        if count%2==1 or count == (len(args.platinumlogos)-1):
            Directory_tex.write(r'''\\
            ''')
        else:
            Directory_tex.write(r'''\hspace{2em}''')
    Directory_tex.write(r'''\end{center}
    ''')
    Directory_tex.write(r'''{\fontspec{Bookman Old Style} \fontsize{16}{19}\selectfont \bf Thank You to Our Gold Corporate Sponsors!}
    \vspace{0em}\begin{center}
    ''')
    for count in range(len(args.goldlogos)):
        logo=args.goldlogos[count]
        max_height=0.3/ceil(len(args.goldlogos)/3)
        Directory_tex.write(r'''\begin{minipage}{0.3\textwidth}\begin{center}\includegraphics[max height='''+str(max_height)+r'''\textheight, max width=\textwidth]{'''+logo+r'''}\vfill\end{center}\end{minipage}
        ''')
        if count%3==2 or count == (len(args.goldlogos)-1):
            Directory_tex.write(r'''\\
            ''')
        else:
            Directory_tex.write(r'''\hspace{1em}''')
    Directory_tex.write(r'''\end{center}
    \newpage
    ''')
    #Back Cover
    Directory_tex.write(r'''\backcover
    ''')
    Directory_tex.write(r'''{\fontspec{Bookman Old Style} \fontsize{16}{19}\selectfont \bf Thank You to Our Directory Sponsors!}\\
    \begin{center}
    \vfill
    ''')
    for logo in args.sponsorlogos:
        Directory_tex.write(r'''\includegraphics[height='''+str(.8/(len(args.sponsorlogos)+1))+r'''\textheight,max width=.85\textwidth]{'''+logo+r'''}\\[3em]
        ''')
    Directory_tex.write(r'''\vfill
    \includegraphics[width=.25\textwidth]{SWE_GEAR}\hspace{3em}
    \includegraphics[width=.25\textwidth]{StackedBlockMwrapped}\hspace{3em}
    \includegraphics[width=.25\textwidth]{RotatedBentwWords}
    ''')
    Directory_tex.write(r'''\end{center}
    ''')
    
    #Close up shop
    Directory_tex.write(tail)
    Directory_tex.close()
    print "Finished writing to: "+args.output
    return company_tex

def compile_directory(company_tex):    
        #This is all the system stuff, all the output has been created already, this is merely
    #to compile the .tex files and
    
    #The actual compile command
    #TODO: Put your own path
    cmd = '/usr/texbin/xelatex -interaction=nonstopmode %(path)s'%{"path":company_tex}

    if True:
        #Clean-up auxiliary files, basically make clean every time
        subprocess.call('rm *.aux', shell=True)
        subprocess.call('rm *.out', shell=True)
        subprocess.call('rm *.log', shell=True)
    
    #Runs the xelatex compilation without giving you all of gibberish output
    print "Compiling "+company_tex+" using XeLaTeX"
    p=subprocess.Popen(cmd.split(), stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    p_data=p.communicate()
    
    #So this is complicated, basically latex compilation doesn't put
    #errors on stderr, only warnings, so if there's an error it prints first
    #the warnings from stderr. The compilation error is buried deep in stdout
    #designated by an exclamation mark, this prints output from stdout in 200
    #character blocks the compilation error and other messages to help diagnose 
    #the error
    if p.returncode!=0:
        print p_data[1]
        error_start=p_data[0].find('!')
        print ''
        print p_data[0][error_start:error_start+200]
        print ''
        while raw_input('Display next 200 characters? (y/n)')=='y':
            error_start+=200
            print ' '
            print p_data[0][error_start:error_start+200]
        exit()
    #If there was no compilation error, compile a second time.
    p=subprocess.Popen(cmd.split(), stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    p_data=p.communicate()
    if p.returncode!=0:
        print p_data[1]
        exit()
    print 'XeLaTeX: Compile Successful'
    
    #Find the output file and copy it to the top level
    file_path = company_tex.split(os.sep)
    file_name = file_path[len(file_path)-1]
    pdf_name=".".join([file_name[:(len(file_name)-4)],"pdf"])
    subprocess.call('cp \"%(file)s\" ..%(sep)sCurrent_Directory.pdf'%{"file":pdf_name,"sep":os.sep},shell=True)
    
def main(argv):
    
    #Setup argument parsing
    parser = argparse.ArgumentParser(description ='Generate the CF Directory')
    parser.add_argument('--chairs', dest='chairs', action='store',
                        default='chairs.csv', help='The csv file containing information on chairs (default: chairs.csv)')
    parser.add_argument('--directors', dest='directors', action='store',
                        default='directors.csv', help='The csv file containing information on directors (default: directors.csv)')
    parser.add_argument('--companies', dest='companies', action='store',
                        default='companies2014-3.csv', help='The csv file containing information on companies (default: companies.csv)')
    parser.add_argument('--output', dest='output', action='store',
                        default='Directory'+str(datetime.now().month)+'_'+str(datetime.now().day)+'.tex', help='The csv file containing information on companies (default: companies.csv)')
    parser.add_argument('--sponsorletters', dest='sponsorletters', action='store', type=lambda x: x.split(','),
                        default=['SponsorLetter2014-1.txt','SponsorLetter2014-2.txt'] , help='A list of directory sponsor letter files (default: 3 files named SponsorLetter#.txt)')
    parser.add_argument('--sponsornames', dest='sponsornames', action='store', nargs='*',
                        default=['Ford','Shell Oil Company'] , help='A list of directory sponsor names (default: 3 probably out of date companies)')
    parser.add_argument('--sponsorlogos', dest='sponsorlogos', action='store', nargs='*',
                        default=['ford_letter_logo','diamond6'] , help='A list of directory sponsor logos (default: [Sponsor1 Sponsor2 Sponsor3])')
    parser.add_argument('--directorletter', dest='directorletter', action='store',
                        default="Director_Welcome2014.txt", help='The text file containing the director welcome letter (default: Director_Welcome.txt)')
    parser.add_argument('--administration', dest='administration', action='store',
                        default="administration.csv", help='The csv file of which administrators to thank (default: administration.csv)')
    parser.add_argument('--ecrc', dest='ecrc', action='store',
                        default="ecrc.csv", help='The csv file of which ECRC people to thank (default: administration.csv)')
    parser.add_argument('--facilities', dest='facilities', action='store',
                        default="facilities.csv", help='The csv file of which facilities people to thank (default: administration.csv)')
    parser.add_argument('--diamondlogos', dest='diamondlogos', action='store', nargs='*',
                        default=['diamond1','diamond2','diamond3','diamond4','diamond5','diamond6','diamond7'] , help='A list of diamond sponsor logo files (default: 6 files named diamond#)')
    parser.add_argument('--platinumlogos', dest='platinumlogos', action='store', nargs='*',
                        default=['platinum1','platinum2','platinum3','platinum4','platinum5','platinum6','platinum7','platinum8'] , help='A list of platinum sponsor logo files (default: 7 files named platinum#)')
    parser.add_argument('--goldlogos', dest='goldlogos', action='store', nargs='*',
                        default=['gold1','gold2','gold3','gold4','gold5','gold6','gold7','gold8','gold9','gold10','gold11'] , help='A list of diamond sponsor logo files (default: 12 files named gold#)')
    
    args = parser.parse_args()    
    #Make sure that there is consistency across the provided directory sponsor information
    if not (len(args.sponsornames)==len(args.sponsorletters) and len(args.sponsornames)==len(args.sponsorlogos)):
        print "ERROR: The list of sponsor names and the list of sponsor letters must be the same length"
        exit(1)
    #Use the provided chairs and directors files to set up information on CF staff
    
    Monday = []
    Tuesday = []
    Monday_Majors = {}
    Tuesday_Majors= {}
    Sponsors = []
    Companies ={"MONDAY":Monday,"TUESDAY":Tuesday}
    
    

    #Setup input and output files and csv readers
    company_spreadsheet_file=args.companies 
    company_spreadsheet=os.path.join("..","Directory_Data",company_spreadsheet_file)
    
    #csv readers
    CompanyReader = unicode_csv_reader(open(company_spreadsheet,'r'), delimiter=',')
    
    #Determine, insofar as possible, which column is which.
    spreadsheet.set_Indices(CompanyReader.next())

    #Process the provided company data.
    for row in CompanyReader:
        major_list = string_cleaner.clean_string_latex(row[spreadsheet.get_Index('Majors')]).split(',')
        major_acron = major_data.major_word2acron(major_list)
        position_list = filter(None,row[spreadsheet.get_Index('Positions')].split(','))
        # degree_list = filter(None,row[spreadsheet.get_Index('Degree')].replace('PhD','Ph.D.').replace('Masters','Master\'s').replace('Bachelors','Bachelor\'s').split(':'))
        degree_list = get_degrees(row[spreadsheet.get_Index('Degree')])
        location_list = filter(None,row[spreadsheet.get_Index('Location')].split(','))
        Generic_Company = {"NAME":string_cleaner.clean_string_latex(row[spreadsheet.get_Index('Name')]),
                           "WEB_NAME":row[spreadsheet.get_Index('Name')],
                           "WEB_PROFILE":row[spreadsheet.get_Index('Profile')],
                           "WEBSITE":row[spreadsheet.get_Index('Website')],
                           "MAJORS":major_list,
                           "MAJOR_ACRON":major_acron,
                           "POSITIONS":position_list,
                           "DEGREES":degree_list,
                           "LOCATION":location_list,
                           "ONLINE-APPLICATION":row[spreadsheet.get_Index('Online-Application')],
                           "HIRING-POLICY":int(get_hiring_policy(row[spreadsheet.get_Index('Hiring-Policy')])),
                           "ATTENDANCE-DATE":get_date(row[spreadsheet.get_Index('Attendance-Date')]),
                           "RECEPTIONS":string_cleaner.YesNo2Bool(row[spreadsheet.get_Index('Receptions')]),
                           "PROFILE":string_cleaner.clean_string_latex(row[spreadsheet.get_Index('Profile')]),
                           "BUILDING":row[spreadsheet.get_Index('Building')].replace('&','\&')}
        #Checks if the company is a sponsor or not, 
        #since this is determined by having any entry in the sponsor column
        # (usually an x), it just checks for existence of
        #any data and also catches the exception if the row is shorter than expected
        try:
            sponsorship = row[spreadsheet.get_Index('Sponsor?')]
            print row[spreadsheet.get_Index('Sponsor?')]
            if(sponsorship in ['Gold','Diamond','Platinum']):
                Generic_Company["SPONSOR"]=True
                Sponsors.append(Generic_Company["NAME"])
            else:
                Generic_Company["SPONSOR"]=False
        except IndexError:
            Generic_Company["SPONSOR"]=False
            
        #Splits the companies by Monday or Tuesday
        #Also populates the companies-by-major containers
        if(Generic_Company['ATTENDANCE-DATE'].lower().find('mon')>-1):
            company_majors = Generic_Company["MAJOR_ACRON"]
            if 'Any' in company_majors:
                
                all_majors = set(major_data.Major_Cipher.values())
                all_majors.remove('Any')
                company_majors = list(all_majors)
                Generic_Company["MAJOR_ACRON"]=company_majors
            for major in company_majors:
                if major in Monday_Majors.keys():
                    Monday_Majors[major].append(Generic_Company["NAME"])
                else:
                    Monday_Majors[major]=[Generic_Company["NAME"]]
            Monday.append(Generic_Company)
        else:
            company_majors = Generic_Company["MAJOR_ACRON"]
            if 'Any' in company_majors:
                all_majors = set(major_data.Major_Cipher.values())
                all_majors.remove('Any')
                company_majors = list(all_majors)
                Generic_Company["MAJOR_ACRON"]=company_majors
            for major in company_majors:
                if major in Tuesday_Majors.keys():
                    Tuesday_Majors[major].append(Generic_Company["NAME"])
                else:
                    Tuesday_Majors[major]=[Generic_Company["NAME"]]
            Tuesday.append(Generic_Company)        

    tex_file = make_directory(Monday,Tuesday,Monday_Majors,Tuesday_Majors,Sponsors,args)
    compile_directory(tex_file)
    make_binder_covers(Monday,Tuesday)
    make_html_page(Monday,Tuesday)
#Calls the main function when run
if __name__ == '__main__':
    sys.exit(main(sys.argv))
    
