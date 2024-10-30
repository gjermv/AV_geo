#! python3
# -*- coding: UTF-8 -*-
# r: pillow


import Rhino
import rhinoscriptsyntax as rs
import scriptcontext as sc
import math
import System
import os
import PIL.Image
import glob
import re
import System.Windows.Forms as forms

import locale
locale.setlocale(locale.LC_ALL, 'nb_NO')

import pandas as pd



def BildemappeTilExcel(bilde_mappe):
    
    bildeFiler = bilde_mappe + '\\*.[jJ][pP]*[gG]' # Burde matche jpg, JPG og jpeg
    bildeOversikt = []
    
    for image in glob.glob(bildeFiler):
        d = {'Pel': 'NA',
             'Pos': 'NA',
             'Side': 'NA',
             'Symbol': 'NA',
             'Bildenavn': 'NA',
             'Kommentar': '',
             'Plassforhold': '',
             'Retning': '',
             'Tiltak': '',
             'Plassering': '',
             'Nedfall': ''}
        
        
        try:
            img = PIL.Image.open(image)
            exif_data = img._getexif()

            head, tail = os.path.split(image)
            d['Bildenavn'] = tail
            
            try:
                image_text = exif_data[270]

                if len(image_text.split('\n')) == 2:

                    if image_text.split('\n')[1].split()[0][1].isdigit():

                        linje0 = image_text.split('\n')[0].split(' ')
                        linje1 = image_text.split('\n')[1][:2]
                        txt = image_text.split('\n')[1][2:].encode('latin1').decode()
                        t = re.sub(r'[^a-zA-Z0-9\s]', '', txt)

                        strings = [linje0[0], linje0[1], linje1, t, tail]
                        
                        d['Side'] = linje0[0]
                        d['Pel'] = linje0[1]
                        d['Symbol'] = linje1
                        d['Kommentar'] = t
                        
                        
                        
                        line_to_copy = '\t'.join(strings)

                        print(line_to_copy)

                    else:
                        linje0 = image_text.split('\n')[0].split(' ')
                        linje1 = image_text.split('\n')[1]
                        # Fix tekst decoding.
                        txt = linje1.encode('latin1').decode()

                        strings = [linje0[0], linje0[1], '-', txt, tail]
                        line_to_copy = '\t'.join(strings)
                        
                        
                        d['Side'] = linje0[0]
                        d['Pel'] = linje0[1]
                        d['Symbol'] = '-'
                        d['Kommentar'] = txt

                        print(line_to_copy)


                elif len(image_text.split('\n')) == 3:
                    linje0 = image_text.split('\n')[0].split(' ')
                    linje1 = image_text.split('\n')[1]
                    linje2 = image_text.split('\n')[2]

                    strings = [linje0[0], linje0[1], linje1, linje2, tail]
                    line_to_copy = '\t'.join(strings)
                    print(line_to_copy)
                    d['Side'] = linje0[0]
                    d['Pel'] = linje0[1]
                    d['Symbol'] = '-'
                    d['Kommentar'] = txt
                
                else:
                    head, tail = os.path.split(image)
                    d['Symbol'] = 'Feil format'
                    d['Kommentar'] = image_text                   
                    
                    print('Error:', image_text)
                    
                bildeOversikt.append(d)

            except:
                d['Symbol'] = 'Bildefeil - Ingen tekst'
                d['Kommentar'] = '-'    
                print('Feil bilde', image)
                bildeOversikt.append(d)
        
        except:
            d['Symbol'] = 'Ikke bilde'
            d['Kommentar'] = '-'
            
  
    return bildeOversikt


def get_directory():
    dialog = forms.FolderBrowserDialog()
    dialog.Description = "Velg mappe med jpg-bilder"
    dialog.ShowNewFolderButton = True

    if dialog.ShowDialog() == forms.DialogResult.OK:
        return dialog.SelectedPath
    return None


if __name__ == '__main__':
    bilde_mappe = get_directory()
    excelFilePath = os.path.join(bilde_mappe,'BildeOversikt.xlsx')

    if os.path.isdir(bilde_mappe):
        res = BildemappeTilExcel(bilde_mappe)
        
        df = pd.DataFrame.from_dict(res)
        
        df.to_excel(excelFilePath)
        
    else:
        print('Filområdet finnes ikke')
