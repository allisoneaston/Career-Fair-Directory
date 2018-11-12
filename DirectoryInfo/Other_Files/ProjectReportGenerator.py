import sys
import csv
import os
import re
from GoogleSpreadSheets import GoogleConnection
projectReportFilePath=r'''ProjectReportData.csv'''
latexFilePath=r'''ProjectReports.tex'''
#regexString=r'''_^(?:(?:https?|ftp)://)(?:\S+(?::\S*)?@)?(?:(?!10(?:\.\d{1,3}){3})(?!127(?:\.\d{1,3}){3})(?!169\.254(?:\.\d{1,3}){2})(?!192\.168(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\x{00a1}-\x{ffff}0-9]+-?)*[a-z\x{00a1}-\x{ffff}0-9]+)(?:\.(?:[a-z\x{00a1}-\x{ffff}0-9]+-?)*[a-z\x{00a1}-\x{ffff}0-9]+)*(?:\.(?:[a-z\x{00a1}-\x{ffff}]{2,})))(?::\d{2,5})?(?:/[^\s]*)?$_iuS'''
#regexString2=r'''_(^|[\s.:;?\-\]<\(])(https?://[-\w;/?:@&=+$\|\_.!~*\|'()\[\]%#,?]+[\w/#](\(\))?)(?=$|[\s',\|\(\).:;?\-\[\]>\)])_i'''
#regexString3=r'''#\b(([\w-]+://?|www[.])[^\s()<>]+(?:\([\w\d]+\)|([^[:punct:]\s]|/)))#iS'''
#regexToken=re.compile(regexString3)


project_name_col		=	1
project_beg_mon_col		=	2
project_beg_day_col		=	3
project_beg_year_col	=	4
project_end_mon_col		=	5
project_end_day_col		=	6
project_end_year_col	=	7
project_new_col			=	8
how_many_electees_col	=	9
how_many_actives_col	=	10
project_leaders_col		=	11
project_description_col	=	12
purpose_objectives_col	=	13
community_col			=	14
university_col			=	15
profession_col			=	16
chapter_col				=	17
honors_col				=	18
other_col				=	19
contact_name_col		=	20
contact_title_col		=	21
contact_number_col		=	22
contact_email_col		=	23
other_info_col			=	24
org_hours_col			=	25
participating_hours_col	=	26
timeframe_col			=	27
other_group_col			=	28
which_group_col			=	29
comments_col			=	30
were_items_needed_col	=	31
what_items_where_col	=	32
cost_col				=	33
problems_col			=	34
recommendations_col		=	35
evaluation_col			=	36
project_exp_col			=	37
best_part_col			=	38
improved_col			=	39
should_continue_col		=	40
participants_col		=	41
relevant_officer_col	=	42
created_officer_files = set()


def is_url(word):
	if word.startswith('http'):
		return True
	if word.startswith('www.'):
		return True
	if word.count('@')>0:
		return False
	if word.endswith('.com') or word.endswith('.edu') or word.endswith('.org'):
		return True
	if word.endswith('.gov') or word.endswith('.html') or word.endswith('.cfm'):
		return True
	if word.endswith('.htm') or word.endswith('.ca') or word.endswith('.net'):
		return True
	return False
def is_email(word):
	if word.count('@')==0:
		return False
	if word.endswith('.com') or word.endswith('.edu') or word.endswith('.org'):
		return True
	if word.endswith('.gov') or word.endswith('.html') or word.endswith('.cfm'):
		return True
	if word.endswith('.htm') or word.endswith('.ca') or word.endswith('.net'):
		return True
	return False
	
def add_page_header(f):
	f.write(r'''\begin{figure}[H]
\centering
\includegraphics{TBP_Project_Report_Header.png}
\end{figure}''')

#Formats phone numbers to be consistent format
#Needs more fleshing out
def process_number(number):
	splitstring=number.rpartition('(')
	number=splitstring[2]
	number=number.replace(')','.')
	number=number.replace('-','.')
	return number

#Cleans up strings that will be put in latex, escaping important characters %,$,&,\ etc
def clean_string_latex(input):
	output=input.replace('\\',r'''\\''')
	output=output.replace('%','\%')
	output=output.replace('$','\$')
	output=output.replace('&','\&')
	output=output.replace('_','\_')
	output=output.replace('#',r'''\#''')
	output_list=output.split()
	count=0
	while count<len(output_list):
		word=output_list[count];
		if is_url(word):
			output_list[count]=r'''\url{'''+word+r'''}'''
		if is_email(word):
			output_list[count]=r'''\href{mailto:'''+word+r'''}{\nolinkurl{'''+word+r'''}}'''
		count=count+1
	output=' '.join(output_list)
	return output
	
projectreport='Final Project Report'
#uncomment below 3 lines to use google sheet, otherwise will use local copy

connect_yes_no=raw_input('Do you want to connect to Google Docs to get the latest sheet? (y/n)')
if(connect_yes_no=='y' or connect_yes_no=='Y'):
	gConnection = GoogleConnection()
	gConnection.ConnectToGoogle()
	gConnection.DownloadSpreadsheet(projectreport,projectReportFilePath)
should_include_pics=raw_input('Do you want to include pictures? (y/n)')
if(should_include_pics=='y' or should_include_pics=='Y'):
	include_pics='x'
else:
		include_pics='n'
	
projectDataReader = csv.reader(open(projectReportFilePath,'r'), delimiter=',')
projectDataReader.next()
latexReport = open(latexFilePath,'w')
historian=raw_input('Enter the preparer\'s name: ')
title=raw_input('Enter the preparer\'s title(i.e. Historian or Secretary): ')
year =raw_input('Enter the school year(i.e. 2011-2012): ')


officer_sheet_header=r'''\documentclass{ProjectReport}
\begin{document}
\begin{titlepage}
\begin{center}
\textsc{\LARGE Tau Beta Pi Project Report Summary}\\[1.5cm]
\textsc{\Large %(year)s}\\[.5cm]
\rule{\linewidth}{0.5mm}\\[.4cm]
{\huge\bfseries %(officer)s}\\[.4cm]
\rule{\linewidth}{0.5mm}\\[1.5cm]
This document contains the project reports related to your officer position. Please keep them as reference as and recommendation.\\
\vfill
{\large Last revised:}\\
{\large \today}
\end{center}
\end{titlepage}

'''

header=r'''\documentclass{ProjectReport}

\begin{document}
\schoolyear{%(year)s}
\historian{%(historian)s}
\preparertitle{%(title)s} 
\maketitle
'''
finished_header=header%{"year":year,"historian":historian,"title":title}
latexReport.write(finished_header)
sectionI=r'''\newproject{%(project_name)s}
\addcontentsline{toc}{section}{%(project_name)s}
\begin{enumerate}[I.]
\item \textbf{Basic Information}
	\begin{enumerate}[A.]
		\item Project Date%(multiple_dates)s\\
			%(project_dates)s
		\item New or Old Project?\\
		$%(new_box)s$ New \hspace{.5in}$%(old_box)s$ Old
		\item Number of Persons who Participated in this Project\\
		Members:~%(active_number)s\hspace{.5in}Electees:~%(electee_number)s
		\item Names of Persons who Participated in this Project\\
		Project Leader%(mult_leaders)s (uniqname)\\
		\begin{tabular}{|l|}\hline %(leaders)s\hline\end{tabular}\paraspace
		\begin{longtable}{|lr|c|r|}\multicolumn{1}{l}{Name}  &\multicolumn{1}{r}{(uniqname)} &
		\multicolumn{1}{c}{Active or Electee}&
		\multicolumn{1}{r}{Number of Hours}\\ \hline
		\endhead
		\hline
		\endfoot
		%(participants)s\end{longtable}
	\end{enumerate}
'''

sectionsII_IV=r'''\item \textbf{General Description of Project}\\
		%(description)s
	\item \textbf{Purpose and Relationship to the Objectives of Michigan Gamma}\\
		%(purpose)s\paraspace
		Who was the Target Audience?\\
		\begin{tabular}{p{1.5in} p{1.5in} p{1.5in}}
			$%(comm_box)s$ Community &$%(U_box)s$  University & $%(prof_box)s$ Profession \\
			$%(chap_box)s$ Chapter & $%(hon_box)s$ Honors/Awards & $%(other_box)s$ Other
		\end{tabular}
	\item \textbf{Organization and Administration}
	\begin{enumerate}[A.] %(contact_info_block)s
		\item Hours Spent on this Project\\
			Organizing: %(org_hours)s\hspace{.3in} Participating: %(part_hours)s
		\item Timeframe from beginning to end of the Project\\
		%(tf1_2)s~1-2 Weeks\hspace{.3in}%(tf3_4)s~3-4 Weeks\hspace{.3in}%(tf5_6)s~5-6 Weeks\hspace{.3in}%(tfmore)s~More\hspace{.3in}
		\item Was this done in conjunction with another group?\hspace{.5in}$%(other_group_yes)s$~Yes\hspace{.5in}$%(other_group_no)s$~No%(other_group_name)s
	\end{enumerate}'''
sectionsV_VII=r'''\item \textbf{Cost and Personnel Requirements}
	\begin{enumerate}[A.] %(general_comments_item)s
		\item Were there any necessary items?\hspace{.5in}$%(items_yes)s$~Yes\hspace{.5in}$%(items_no)s$~No %(necessary_items)s
		\item What was the total cost associated with this project?\\
		%(cost)s
	\end{enumerate}
	\item \textbf{Special Problems}\\
	%(special_problems)s
	\item \textbf{Recommendations}\\
	%(recommendations)s
	'''
sectionVIII=r'''\item\textbf{Overall Evaluation}
	\begin{enumerate}[A.]  %(eval_gen_comments)s
	\item Overall Project Experience \\
		Best
		\begin{tabular}{p{1cm} p{1cm} p{1cm} p{1cm} p{1cm}}
			1&2&3&4&5\\
			$%(one_box)s$ & $%(two_box)s$ & $%(three_box)s$ & $%(four_box)s$ & $%(five_box)s$
		\end{tabular}
		Worst
	\item What was the best part about this project?\\
		%(best_part)s
	\item What could be improved about this project?\\
		%(improved)s
	\item Would you recommend continuing this project?\\
		%(continue)s
	\end{enumerate}'''
end_sections=r'''\end{enumerate}'''
for row in projectDataReader:
	if(row[0].replace(' ','')=='Divider'):
		if(row[1]=='Category'):
			latexReport.write(r'''\part{%s}
			'''%(row[2]))
		elif(row[1]=='Semester'):
			latexReport.write(r'''\semester{%s}
			'''%(row[2]))
		continue
	officer_name=row[relevant_officer_col]	
	files_dir='./%s'%(officer_name)
	officer_sheet_path="%s.tex"%(files_dir)
	if not os.path.exists(files_dir):
		os.mkdir(files_dir)
	latex_sheet_path="%s/%s.tex"%(files_dir,row[43])
	latex_sheet=open(latex_sheet_path,'w')
	if officer_name not in created_officer_files:
		created_officer_files.add(officer_name)
		officer_sheet=open(officer_sheet_path,'w')
		officer_sheet.write(officer_sheet_header%{"year":year,"officer":officer_name})
		officer_sheet.close()
	
	add_page_header(latex_sheet)
	
	#Name and dates
	project_name=row[project_name_col]
	multiple_dates='s'
	b_month=row[project_beg_mon_col]
	b_day=row[project_beg_day_col]
	b_year=row[project_beg_year_col]
	e_month=row[project_end_mon_col]
	e_day=row[project_end_day_col]
	e_year=row[project_end_year_col]
	if(b_month == e_month and b_day == e_day and b_year == e_year):
		multiple_dates=''
		project_dates = b_month+" "+b_day+", "+b_year
	else:
		project_dates = b_month+" "+b_day+", "+b_year+" - "+e_month+" "+e_day+", "+e_year
		
	#New or Old
	if(row[project_new_col]=="Yes"):
		new_box=r'''\boxtimes'''
		old_box=r'''\Box'''
	else:
		old_box=r'''\boxtimes'''
		new_box=r'''\Box'''
		
	#Leaders and Participants
	leader_names=row[project_leaders_col]
	mult_leaders=''
	leader_names=leader_names.replace(')',r''')\\''')
	participants=row[participants_col]
	participants=participants.replace('|','&')
	participants=participants.replace('\\','&')
	participants=participants.replace('(','& (')
	participants=participants.replace('\n',r'''\\''')
	if(leader_names.count(')')>1):
		mult_leaders='s'

	#Purpose Checkboxes
	comm_box=r'''\Box'''
	U_box=r'''\Box'''
	prof_box=r'''\Box'''
	chap_box=r'''\Box'''
	hon_box=r'''\Box'''
	other_box=r'''\Box'''
	if(row[community_col]=='X'):
		comm_box=r'''\boxtimes'''
	if(row[university_col]=='X'):
		U_box=r'''\boxtimes'''
	if(row[profession_col]=='X'):
		prof_box=r'''\boxtimes'''
	if(row[chapter_col]=='X'):
		chap_box=r'''\boxtimes'''
	if(row[honors_col]=='X'):
		hon_box=r'''\boxtimes'''
	if(row[other_col]=='X'):
		other_box=r'''\boxtimes'''
		

	contact_name=row[contact_name_col]
	contact_title=clean_string_latex(row[contact_title_col])
	contact_number=clean_string_latex(row[contact_number_col])
	contact_email=clean_string_latex(row[contact_email_col])
	contact_other_info=clean_string_latex(row[other_info_col])
	contact_section_full=''
	name_insert=''
	title_insert=''
	phone_insert=''
	email_insert=''
	other_info_insert=''
	contact_info_block=''
	if(contact_name!='' or contact_title!='' or contact_number !='' or contact_email!='' or contact_other_info!=''):
		if(contact_name!=''):
			name_insert=r'''Name: 	&	%s\\'''%(contact_name)
		if(contact_title!=''):
			title_insert=r'''Title: 	&	%s\\'''%(contact_title)
		if(contact_number!=''):
			phone_insert=r'''Phone \#: 	&	%s\\'''%(process_number(contact_number))
		if(contact_email!=''):
			email_insert=r'''Email: 	&	%s\\'''%(contact_email)
		if(contact_other_info!=''):
			other_info_insert=r'''Other Info: 	&	%s\\'''%(contact_other_info)
		contact_info_block=r'''
		\item Contact Information\\
			\begin{tabular}{l p{5in}}
				%s
				%s
				%s
				%s
				%s
			\end{tabular}'''%(name_insert,title_insert,phone_insert,email_insert,other_info_insert)
	
	#TimeFrame 
	timeFrame1_2=r'''$\Box$'''
	timeFrame3_4=r'''$\Box$'''
	timeFrame5_6=r'''$\Box$'''
	timeFrameMore=r'''$\Box$'''
	if(row[timeframe_col]=='1-2 weeks'):
		timeFrame1_2=r'''$\boxtimes$'''
	elif(row[timeframe_col]=='3-4 weeks'):
		timeFrame3_4=r'''$\boxtimes$'''
	elif(row[timeframe_col]=='5-6 weeks'):
		timeFrame5_6=r'''$\boxtimes$'''
	else:
		timeFrameMore=r'''$\boxtimes$'''
	#Other Group involved?	
	if(row[other_group_col]=='Yes'):
		other_group_yes=r'''\boxtimes'''
		other_group_no=r'''\Box'''
		other_group_name=r'''\\
		With whom?\hspace{.5in} %s'''%(row[which_group_col])
	else:
		other_group_yes=r'''\Box'''
		other_group_no=r'''\boxtimes'''
		other_group_name=''
	
	#Cost and Reqs
	gen_comments=''
	if(row[comments_col]!=''):
		gen_comments=r'''
		\item General Comments\\ %s'''%(clean_string_latex(row[comments_col]))
	if(row[were_items_needed_col]=='Yes'):
		items_yes=r'''\boxtimes'''
		items_no=r'''\Box'''
		necessary_items=r'''\\
		What were they and where did you get them?\\ %s'''%(clean_string_latex((row[what_items_where_col]).replace('\\','/')))
	else:
		items_yes=r'''\Box'''
		items_no=r'''\boxtimes'''
		necessary_items=''
	cost=row[cost_col]
	if(cost.count('$')==0):
		cost='\$'+cost
	else:
		cost=cost.replace('$','\$')
	
	general_eval=''
	one_box=r'''\Box'''
	two_box=r'''\Box'''
	three_box=r'''\Box'''
	four_box=r'''\Box'''
	five_box=r'''\Box'''
	if(row[evaluation_col]!=''):
		general_eval=r'''
			\item General Comments\\
			%s'''%(clean_string_latex(row[evaluation_col]))
	if(row[project_exp_col]=='1'):
		one_box=r'''\boxtimes'''
	if(row[project_exp_col]=='2'):
		two_box=r'''\boxtimes'''
	if(row[project_exp_col]=='3'):
		three_box=r'''\boxtimes'''
	if(row[project_exp_col]=='4'):
		four_box=r'''\boxtimes'''
	if(row[project_exp_col]=='5'):
		five_box=r'''\boxtimes'''
	section_I_full=sectionI%{"project_name":clean_string_latex(project_name),"multiple_dates":multiple_dates,
							"project_dates":project_dates,"new_box":new_box,"old_box":old_box,"active_number":row[how_many_actives_col],
							"electee_number":row[how_many_electees_col],"leaders":leader_names,
							"mult_leaders":mult_leaders,"participants":participants}
	section_II_IV_full=sectionsII_IV%{"description":clean_string_latex(row[project_description_col]),"purpose":row[purpose_objectives_col],"comm_box":comm_box,"U_box":U_box,"chap_box":chap_box,
							"prof_box":prof_box,"other_box":other_box,"hon_box":hon_box,"contact_info_block":contact_info_block,"part_hours":row[participating_hours_col],"org_hours":row[org_hours_col],
							"other_group_yes":other_group_yes,"other_group_no":other_group_no,"other_group_name":other_group_name,"tf1_2":timeFrame1_2,"tf3_4":timeFrame3_4,"tf5_6":timeFrame5_6,"tfmore":timeFrameMore}
	
	section_V_VII_full =sectionsV_VII%{"general_comments_item":gen_comments,"items_yes":items_yes,"items_no":items_no,"necessary_items":necessary_items,"cost":cost,
							"special_problems":row[problems_col],"recommendations":clean_string_latex(row[recommendations_col])}
	sectionVIII_full = sectionVIII%{"eval_gen_comments":general_eval,"one_box":one_box,"two_box":two_box,"three_box":three_box,"four_box":four_box,"five_box":five_box,
							"best_part":row[best_part_col], "improved": clean_string_latex(row[improved_col]), "continue": row[should_continue_col]}
	latex_sheet.write(section_I_full)
	latex_sheet.write(section_II_IV_full)
	latex_sheet.write(section_V_VII_full)
	latex_sheet.write(sectionVIII_full)
	pic_check=include_pics#commented out--to skip pictures put n for picture promt put x
	while(pic_check!='y' and pic_check!='n'):
		pic_check=raw_input("Are there pictures for "+project_name+"? (y/n)")
	if(pic_check=='y'):
		pic_count=1
		pic_name='none'
		while(pic_name!='q'):
			pic_name=raw_input("Enter the file name for picture #%d (q if this event has no more pictures): "%(pic_count))
			try:
				temp_f=open(pic_name)
				temp_f.close()
				pic_caption=raw_input("Enter a caption: ")
				if(pic_count==1):
					latex_sheet.write(r'''\item \textbf{Pictures}''')
				pic_count=pic_count+1
				latex_sheet.write(r'''\begin{figure}[H]
						\centering
						\caption{%s}
						\label{%s}
						\includegraphics[width=5.4in]{%s}
					\end{figure}'''%(pic_caption,"%s%s"%(project_name,pic_count),pic_name))
			except IOError as e:
				if(pic_name!='q'):
					print "Invalid File Name"
					
	latex_sheet.write(end_sections)	
	latex_sheet.close()
	latexReport.write(r'''\input{%s}
	'''%latex_sheet_path)
	latexReport.write(r'''\newpage
	''')
	officer_sheet=open(officer_sheet_path,'a')
	officer_sheet.write(r'''\input{%s}
	'''%latex_sheet_path)
	officer_sheet.write(r'''\newpage
	''')
	officer_sheet.close()
	
	
for pos in created_officer_files:
	officer_sheet_finish=open("./%s.tex"%(pos),'a')
	officer_sheet_finish.write(r'''\end{document}''')
	officer_sheet_finish.close()
	cmd = 'pdflatex \"%(path)s\"'%{"path":"./%s.tex"%(pos)}
	os.system(cmd)
	

latexReport.write(r'''\end{document}''')
latexReport.close()

cmd = 'pdflatex \"%(path)s\"'%{"path":latexFilePath}
os.system(cmd)