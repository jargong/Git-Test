from docx import *
import os
import re

indir = "D:\\Files\\Documents\\Downloads\\" #put in the directory you are getting the docx files from. Remember to use \\ instead of \
odir = "D:\\Files\\Documents\\Downloads\\" #put in the directory you want to spit the txt files out to.
banned_words = [""] #strings in brackets will be removed. you can put multiple words in here, separated by commas. Ex. ["poop","ass"]
banned_phrases = [""] #same as above, just put in instances of multiple words.
bracketHate = False #True: Removes the WHOLE word if ANY part of a bracket is found.


for file in os.listdir(indir):
	paragraph_list = []
	if file.endswith(".docx"):
		
		fname = file.split(".")[0]
		document = Document(indir + fname + ".docx")
		for para in document.paragraphs:
			paragraph_list.append(para.text.encode('ascii', 'ignore'))
		new_txt_file = open(odir + fname + ".txt" ,"w")

		for index_paralist,paragraph in enumerate(paragraph_list):

			para_temp = paragraph
			
			#removes brackets
			if bracketHate == True:
				para_temp = re.sub("\([\s\S]*\)", "", para_temp)
				para_temp = re.sub("\[[\s\S]*\]", "", para_temp)

			#removes banned phrases
			for phrase in banned_phrases:
				if phrase.lower() in para_temp.lower():
					para_temp = para_temp.replace(phrase, "")
			para_split = para_temp.split(" ")
			para_split = filter(None, para_split)

			#removes banned words
			for i, split_word in enumerate(para_split):
				for bword in banned_words:
					if bword.lower() == split_word.lower():
						para_split.remove(split_word)
				if "\n" in split_word:
					para_split[i]=split_word.strip("\n")
			#puts the list back together
			para_join = " ".join(para_split)
			paragraph_list[index_paralist] = para_join
		print "All unecessary words removed!"


		#deletes everything before 
		useful_info_start = 0
		keep_going = True
		for index_paralist,paragraph in enumerate(paragraph_list):
			if keep_going == True:
				if re.search(".*:",paragraph):
					useful_info_start = index_paralist
					keep_going = False
		del paragraph_list[0:useful_info_start]

		print paragraph_list
		for line in paragraph_list:
			if not line == "":
				new_txt_file.write(line.encode('ascii', 'ignore'))
				new_txt_file.write("\n")
		print fname + ".txt" + " finished!"