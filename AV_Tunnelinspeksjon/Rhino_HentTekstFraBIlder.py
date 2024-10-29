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



def BildemappeTilExcel(bilde_mappe):
    for image in glob.glob(bilde_mappe):
        try:
            img = PIL.Image.open(image)
            exif_data = img._getexif()

            head, tail = os.path.split(image)

            try:
                image_text = exif_data[270]

                if len(image_text.split('\n')) == 2:

                    if image_text.split('\n')[1].split()[0][1].isdigit():

                        linje0 = image_text.split('\n')[0].split(' ')
                        linje1 = image_text.split('\n')[1][:2]
                        txt = image_text.split('\n')[1][2:].encode('latin1').decode()
                        t = re.sub(r'[^a-zA-Z0-9\s]', '', txt)

                        strings = [linje0[0], linje0[1], linje1, t, tail]
                        line_to_copy = '\t'.join(strings)

                        print(line_to_copy)

                    else:
                        linje0 = image_text.split('\n')[0].split(' ')
                        linje1 = image_text.split('\n')[1]
                        # Fix tekst decoding.
                        txt = linje1.encode('latin1').decode()

                        strings = [linje0[0], linje0[1], '-', txt, tail]
                        line_to_copy = '\t'.join(strings)

                        print(line_to_copy)


                elif len(image_text.split('\n')) == 3:
                    linje0 = image_text.split('\n')[0].split(' ')
                    linje1 = image_text.split('\n')[1]
                    linje2 = image_text.split('\n')[2]

                    strings = [linje0[0], linje0[1], linje1, linje2, tail]
                    line_to_copy = '\t'.join(strings)
                    print(line_to_copy)

                else:
                    head, tail = os.path.split(image)
                    print('Error:', image_text)


            except:
                print('Feil bilde', image)
        except:
            print('Ikke et bilde', image)


def get_directory():
    dialog = forms.FolderBrowserDialog()
    dialog.Description = "Velg mappen med bilder"
    dialog.ShowNewFolderButton = True

    if dialog.ShowDialog() == forms.DialogResult.OK:
        return dialog.SelectedPath
    return None


if __name__ == '__main__':
    bilde_mappe = get_directory()
    print(bilde_mappe)
    if os.path.isdir(bilde_mappe):
        bilde_mappe = f'{bilde_mappe}\\*'
        print(bilde_mappe)
        BildemappeTilExcel(bilde_mappe)
    else:
        print('Filområdet finnes ikke')
