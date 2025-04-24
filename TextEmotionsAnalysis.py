import nrclex
from nrclex import NRCLex

import nltk
#nltk.download()

import numpy as np
import pandas as pd
import os
import sys

from imutils import paths
import argparse
import numpy as np
import os
import sys
import cv2
import shutil
import re

import matplotlib.pyplot as plt 

import glob

import docx 
from docx.enum.text import WD_COLOR_INDEX 

from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE



#==============================================================================

parser = argparse.ArgumentParser()
parser.add_argument("--Path",type=str,default="None",help='Path.')
parser.add_argument("--QRebuild", type=int, default=0,help='QRebuild.')
args = parser.parse_args()


currentDirectory = os.getcwd()
PathW=os.path.dirname(sys.argv[0])
QView=False

print("CurrentDirectory="+currentDirectory)
print("PathW="+PathW)

PathDataSet=args.Path
if (PathDataSet=='None'):
    PathDataSet=currentDirectory+"\DATA"
    print("Path dataset="+PathDataSet)


ListFiles=glob.glob(PathDataSet+"/*.txt")

if QView:
    print(ListFiles)



def GetCoeffEmotion(ch):
    coeff=0;
    Content=str(ch)
    if (len(ch)>0):
        if (Content.find("fear")>0):
            coeff=coeff-1 
        if (Content.find("anger")>0):
            coeff=coeff-1 
        if (Content.find("anticip")>0):
            coeff=coeff+1 
        if (Content.find("trust")>0):
            coeff=coeff+1 
        if (Content.find("surprise")>0):
            coeff=coeff+1  
        if (Content.find("positive")>0):
           coeff=coeff+1  
        if (Content.find("negative")>0):
           coeff=coeff-1   
        if (Content.find("sadness")>0):
            coeff=coeff-1  
        if (Content.find("disgust")>0):
            coeff=coeff-1    
        if (Content.find("joy")>0):
            coeff=coeff+1   
    return(coeff)

for i in range(0,len(ListFiles)):
    file1 = open(ListFiles[i])
    TextToRead=file1.read()
    file1.close()
    
    if (1==1):
        if QView:
            print("");
            print("");
            print(TextToRead)
            print("");
            print("");
        
        FileName=ListFiles[i][:-4]
        
        currentDirectory = os.getcwd()
        #print(currentDirectory)
        
        
        #Instantiate text object (for best results, 'text' should be unicode).
        
        #text_object = NRCLex('text')
        
        text_object = NRCLex(TextToRead)
        
        
        #List Sentences
        #print("NbSentences="+str(len(text_object.sentences)))
        
        
        #FICH1 = open(FileName+"_Repport_Sentences.csv", "w")
        for i in range(len(text_object.sentences)):
            Phrase=text_object.sentences[i]
            Phrase_object = NRCLex(str(Phrase))
            if QView:
                print(Phrase)
                print(Phrase_object.top_emotions)
                print(GetCoeffEmotion(Phrase_object.top_emotions))
                print("")

        
        
#1   emotion.words 	Return words list.
#2	emotion.sentences	Return sentences list.
#3	emotion.affect_list	Return affect list.
#4	emotion.affect_dict	Return affect dictionary.
#5	emotion.raw_emotion_scores	Return raw emotional counts.
#6	emotion.top_emotions	Return highest emotions.
#7	emotion.affect_frequencies

        #print("=Return Top Emotions=================================================================================")
        #print(text_object.top_emotions)
        
        #Return affect dictionary.
        #print("=Return affect dictionary=================================================================================")
        Name=tuple(text_object.affect_dict)
        #Name
        #text_object.affect_dict[Name[0]]
        
        FICH1 = open(FileName+"_Report_Analysis.csv", "w")
        for i in range(len(Name)):
            if QView:
                print(Name[i]+" "+str(text_object.affect_dict[Name[i]]))
                print("")
            Sch=Name[i]+","+str(text_object.affect_dict[Name[i]])
            FICH1.write(Sch)
            FICH1.write("\n")
        FICH1.close()
        
        #Return affect frequencies.
        FICH2 = open(FileName+"_Report.csv", "w")
        FICH2.write("Emotions,Frequencies\n")
        if QView:
            print("")
            print("Nb frequencies="+str(len(text_object.affect_frequencies)))
        Name=tuple(text_object.affect_frequencies)
        for i in range(len(text_object.affect_frequencies)):
            if QView:
                print(Name[i]+" "+str(text_object.affect_frequencies[Name[i]]))
            Sch=Name[i]+","+str(text_object.affect_frequencies[Name[i]])
            FICH2.write(Sch)
            FICH2.write("\n")
       
        if QView:
            print("===========================================================================")
            print("***************************************************************************")
        FICH2.close()
        
        data = pd.read_csv(FileName+"_Report.csv") 
        df = pd.DataFrame(data) 
      
        X = list(df.iloc[:, 0]) 
        Y = list(df.iloc[:, 1]) 
      
        fig = plt.figure() 
        max_y_lim = max(Y)+0.01
        min_y_lim = min(Y)
        plt.ylim(min_y_lim, max_y_lim)

        #plt.bar(X, Y, color='g') 
        bars=plt.bar(X, Y) 
        bars[0].set_color('red') #Fear
        bars[1].set_color('red') #Anger
        bars[2].set_color('gray') #Disgust
        bars[3].set_color('blue') #Trust
        bars[4].set_color('pink') #Surprise
        bars[5].set_color('green') #Positive
        bars[6].set_color('red') #Negative
        bars[7].set_color('red') #Sadness
        bars[8].set_color('red') #Disgust
        bars[9].set_color('yellow') #Joy


        plt.title("Analysis of text emotions") 
        plt.xlabel("") 
        plt.ylabel("Percentage") 
        plt.xticks(rotation=45)
        plt.savefig(FileName+"_Report.jpg")
        
        #fig2 = plt.figure() 
        #plt.pie(Y, labels = X, startangle = 90)
        #plt.savefig(FileName+"_Repport_Pie.jpg")
        #plt.show() 
    
    if (1==1):
    
        # Create an instance of a word document 
        doc = docx.Document() 
        
        if QView:
            print("NbSentences="+str(len(text_object.sentences)))
          
        # Add a Title to the document  
        doc.add_heading('Analysis of text emotions', 0) 
        para = doc.add_paragraph("") 
        
        font_styles = doc.styles
        font_charstyle = font_styles.add_style('CommentsStyle', WD_STYLE_TYPE.CHARACTER)
        font_object = font_charstyle.font
        font_object.size = Pt(10)
        font_object.name = 'Times New Roman'
        
           
        for i in range(len(text_object.sentences)):
            Phrase=str(text_object.sentences[i])+" "
            Phrase_object = NRCLex(Phrase)
            Content=str(Phrase_object.top_emotions)
            if (len(Phrase_object.top_emotions)==1):
                if (Content.find("fear")>0):
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.RED
                elif (Content.find("anger")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.DARK_RED
                elif (Content.find("anticip")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.GRAY_25
                elif (Content.find("trust")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.TURQUOISE
                elif (Content.find("surprise")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.PINK
                elif (Content.find("positive")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN 
                elif (Content.find("negative")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.RED 
                elif (Content.find("sadness")>0):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.VIOLET 
                elif (Content.find("disgust")>0):
                      para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.DARK_YELLOW
                elif (Content.find("joy")>0):
                      para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.YELLOW 
                else:
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.WHITE 
            else:
                para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.WHITE
          
          
        # Now save the document to a location  
        doc.save(FileName+"_Report.docx")
    
    
    if (1==0):
    
        # Create an instance of a word document 
        doc = docx.Document() 
        
        if QView:
            print("NbSentences="+str(len(text_object.sentences)))
          
        # Add a Title to the document  
        doc.add_heading('Analysis of text emotions', 0) 
        para = doc.add_paragraph("") 
        
        font_styles = doc.styles
        font_charstyle = font_styles.add_style('CommentsStyle', WD_STYLE_TYPE.CHARACTER)
        font_object = font_charstyle.font
        font_object.size = Pt(10)
        font_object.name = 'Times New Roman'
        
        
        for i in range(len(text_object.sentences)):
            Phrase=str(text_object.sentences[i])+" "
            Phrase_object = NRCLex(Phrase)
            
            for j in range(len(Phrase_object.words)):
                Phrase=str(Phrase_object.words[j])+" "
    
                Content=str(Phrase_object.top_emotions)
                if (len(Phrase_object.top_emotions)==1):
                    if (Content.find("fear")>0):
                        para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.RED
                    elif (Content.find("anger")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.DARK_RED
                    elif (Content.find("anticip")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.GRAY_25
                    elif (Content.find("trust")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.TURQUOISE
                    elif (Content.find("surprise")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.PINK
                    elif (Content.find("positive")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN 
                    elif (Content.find("negative")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.RED 
                    elif (Content.find("sadness")>0):
                         para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.VIOLET 
                    elif (Content.find("disgust")>0):
                          para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.DARK_YELLOW
                    elif (Content.find("joy")>0):
                          para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.YELLOW 
                    else:
                        para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.WHITE 
                else:
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.WHITE
              
          
        # Now save the document to a location  
        doc.save(FileName+"_Report_zoom.docx")
    
    
    
    
    if (1==1):
    
        # Create an instance of a word document 
        doc = docx.Document() 
        
        if QView:
            print("NbSentences="+str(len(text_object.sentences)))
          
        # Add a Title to the document  
        doc.add_heading('Analysis of text emotions', 0) 
        para = doc.add_paragraph("") 
        
        font_styles = doc.styles
        font_charstyle = font_styles.add_style('CommentsStyle', WD_STYLE_TYPE.CHARACTER)
        font_object = font_charstyle.font
        font_object.size = Pt(10)
        font_object.name = 'Times New Roman'
        
           
        for i in range(len(text_object.sentences)):
            Phrase=str(text_object.sentences[i])+" "
            Phrase_object = NRCLex(Phrase)
            Num=GetCoeffEmotion(Phrase_object.top_emotions)
            
            if (len(Phrase_object.top_emotions)>0):
                if (Num<=-4):
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.DARK_RED
                elif (Num<=-3):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.RED
                elif (Num<=-1):
                     para.add_run(Phrase, style='CommentsStyle').font.highlight_color = WD_COLOR_INDEX.VIOLET
                elif (Num==0):
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.WHITE
                elif (Num>=4):
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.BRIGHT_GREEN
                elif (Num>=3):
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.GREEN
                elif (Num>=1):
                    para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.GRAY_25
               
            else:
                para.add_run(Phrase, style='CommentsStyle').font.highlight_color =WD_COLOR_INDEX.WHITE
                
          
        # Now save the document to a location  
        doc.save(FileName+"_Report_Colors.docx")
        
        

if (1==0):
    
    # Create an instance of a word document 
    doc = docx.Document() 
      
    # Add a Title to the document  
    doc.add_heading('GeeksForGeeks', 0) 
      
    # Adding Auto Styled Highlighted paragraph 
    doc.add_heading('AUTO Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.AUTO 
      
    # Adding Black Styled Highlighted paragraph 
    doc.add_heading('BLACK Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.BLACK 
      
    # Adding Blue Styled Highlighted paragraph 
    doc.add_heading('BLUE Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.BLUE 
      
    # Adding Bright Green Styled Highlighted paragraph 
    doc.add_heading('BRIGHT_GREEN Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN 
      
    # Adding Dark Blue Styled Highlighted paragraph 
    doc.add_heading('DARK_BLUE Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.DARK_BLUE 
      
    # Adding Dark Red Styled Highlighted paragraph 
    doc.add_heading('DARK_RED Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.DARK_RED 
      
    # Adding Dark Yellow Styled Highlighted paragraph 
    doc.add_heading('DARK_YELLOW Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.DARK_YELLOW 
      
    # Adding GRAY25 Styled Highlighted paragraph 
    doc.add_heading('GRAY_25 Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.GRAY_25 
      
    # Adding GRAY50 Styled Highlighted paragraph 
    doc.add_heading('GRAY_50 Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.GRAY_50 
      
    # Adding GREEN Styled Highlighted paragraph 
    doc.add_heading('GREEN Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.GREEN 
      
    # Adding Pink Styled Highlighted paragraph 
    doc.add_heading('PINK Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.PINK 
      
    # Adding Red Styled Highlighted paragraph 
    doc.add_heading('RED Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.RED 
      
    # Adding Teal Styled Highlighted paragraph 
    doc.add_heading('TEAL Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.TEAL 
      
    # Adding Turquoise Styled Highlighted paragraph 
    doc.add_heading('TURQUOISE Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.TURQUOISE 
      
    # Adding Violet Styled Highlighted paragraph 
    doc.add_heading('VIOLET Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.VIOLET 
      
    # Adding White Styled Highlighted paragraph 
    doc.add_heading('WHITE Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.WHITE 
      
    # Adding Yellow Styled Highlighted paragraph 
    doc.add_heading('YELLOW Style:', 3) 
    doc.add_paragraph().add_run('GeeksforGeeks is a Computer Science portal for geeks.'
                      ).font.highlight_color = WD_COLOR_INDEX.YELLOW 
      
    # Now save the document to a location  
    doc.save('Color.docx')
    

