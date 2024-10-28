#! python3
# -*- coding: UTF-8 -*-
# r: pandas, openpyxl


import Rhino
import rhinoscriptsyntax as rs
import scriptcontext as sc
import math
import System

import openpyxl
import os

def tegneforklaring():
    tf = dict()
    tf['Berg (F)'] = '_'
    tf['F1'] = 'Nedfall d<0.3m'
    tf['F2'] = 'Nedfall d>0.3m'
    tf['F3'] = 'Avløste blokker'
    tf['F4'] = 'Bom'
    tf['F5'] = 'Avskalling og bergslag'
    tf['F6'] = 'Utpressing'
    tf['F7'] = 'Vanninntrengning'
    tf['F8'] = 'Iskjøving'
    tf['Sprøytebetong (S)'] = ' '
    tf['S1'] = 'Nedfall'
    tf['S2'] = 'Riss'
    tf['S3'] = 'Sprekker'
    tf['S4'] = ['Bom']
    tf['S5'] = 'Avskalling'
    tf['S6'] = 'Utpressing'
    tf['S7'] = 'Vanninntrengning'
    tf['S8'] = 'Iskjøving'
    tf['S9'] = 'Nedbrytning (vannkjemi, bakterier)'
    tf['Bolter til bergsikring (B)'] = ' '
    tf['B1A-E'] = 'Korrosjon, Rustgrad A-E'
    tf['B2'] = 'Vrakbolt'
    tf['B3'] = 'Utpressing'
    tf['B4'] = 'Deformasjon'
    tf['Øvrige skader/mangler (M)'] = ' '
    tf['M1'] = 'Manglende / ikke utført bergsikring'
    tf['M2'] = 'Mangler ved utført bergsikring'
    tf['M3'] = 'Manglende vedlikeholdsrensk'
    tf['M4'] = 'Skader på vann-/ frostsikringshvelv'
    tf['Framkommelighet'] = ' '
    tf['X'] = ' Stengt'
    tf['I/X'] = 'Trangt'
    tf[' '] = ' '
    tf['Anbefalte tiltak'] = ' '
    tf['Re'] = 'Rensk'
    tf['B'] = 'Bolt'
    tf['FB'] = 'Fjellbånd'
    tf['SFR'] = 'Sprøytebetong'
    tf['V'] = 'Vann/frostsikring'
    tf['A'] = 'Annet'
    
    return tf

def dwgLag():
    lag = {
        'Vann-og frostsikring_skadereg': (255, 0, 0),
        'Manglende utført bergsikring': (43, 23, 200),
        'Forsagere': (0, 182, 25),
        'Berg_skadereg': (0, 255, 0),
        'Fremkommelighet': (21, 25, 34),
        'Legend': (0, 0, 0)
    }

    return lag

def lagTunnellinjer(start, stop):
    rs.AddLayer("Kartleggingsskjema")
    rs.CurrentLayer("Kartleggingsskjema")

    T_scale = 200
    vl_max = 110/500*T_scale 
    vl_mid = 50/500*T_scale 
    hPos = [-vl_max,-vl_mid, 0, vl_mid, vl_max]  
    
    for h in hPos:
        rs.AddLine((h,start),(h,stop+20))

    for v in range(start, stop+30, 20):
        rs.AddLine((-vl_max,v),(vl_max,v))
        rs.AddText(str(v),(-vl_max-5,v), 2, font=None, font_style=0, justification=131072+4)


def CreateOverskift(tunnelNavn):

    tabellBredde = 180
    tabellHoyde = 8
    linjeA = 80
    linjeB = 110
    textBuffer = 1
    

    rs.CurrentLayer('Default')
    rs.AddText('Vedlegg 1 - Kartleggingsskjema', (textBuffer, 13), 3)

    rs.AddText('Tunnel (og løp)', (textBuffer, 7.5), 1.5) 
    rs.AddText('Dato', (linjeA+textBuffer, 7.5), 1.5) 
    rs.AddText('Registert av', (linjeB+textBuffer, 7.5), 1.5) 
    
    rs.AddText(tunnelNavn, (textBuffer, 4), 3)

    rs.AddLine((0, tabellHoyde), (tabellBredde, tabellHoyde))
    rs.AddLine((0, 0), (tabellBredde, 0))
    rs.AddLine((0, 0), (0, tabellHoyde))
    rs.AddLine((linjeA, 0), (linjeA, tabellHoyde))
    rs.AddLine((linjeB, 0), (linjeB, tabellHoyde))
    rs.AddLine((tabellBredde, 0), (tabellBredde, tabellHoyde))

    objrefs = sc.doc.Objects.FindByLayer('Default')
    base_point = Rhino.Geometry.Point3d(0,0,0)
    idef_name = 'Blokk_Overskrift'
    
    geometry = [item.Geometry for item in objrefs]
    attributes = [item.Attributes for item in objrefs]

    idef_index = sc.doc.InstanceDefinitions.Add(idef_name, "", base_point, geometry, attributes)
    rs.DeleteObjects(objrefs)


def CreateLegend():
    rs.CurrentLayer('Legend')
    legend = tegneforklaring()
    radHoyde = 4
    tabellBredde = 65
    tabellHoyde = (len(legend)) * radHoyde
    skriftStr = 2
    tabA = 2
    tabB = 12

    for i, key in enumerate(legend):
        print(key, i)
        rs.AddLine((0, 0 - (i * radHoyde)), (tabellBredde, 0 - (i * radHoyde)))
        rs.AddText(key, (tabA, -0.85 - (i * radHoyde)), skriftStr)
        rs.AddText(legend[key], (tabB, -0.85 - (i * radHoyde)), skriftStr)
    rs.AddLine((0, tabellHoyde+radHoyde), (tabellBredde, tabellHoyde+radHoyde))  
    rs.AddLine((0, tabellHoyde+radHoyde*2), (tabellBredde, tabellHoyde+radHoyde))   

    rs.AddLine((0, 0), (tabellBredde, 0))
    rs.AddLine((0, 0), (0, -tabellHoyde))
    rs.AddLine((tabellBredde, 0), (tabellBredde, -tabellHoyde))

    rs.AddText('Kommentarer:', (tabA, -170), 2)
    rs.AddLine((0, -168), (tabellBredde, -168))
    rs.AddLine((0, -174), (tabellBredde, -174))
    rs.AddLine((0, -266), (tabellBredde, -266))
    rs.AddLine((0, -168), (0, -266))
    rs.AddLine((tabellBredde, -168), (tabellBredde, -266))

    objrefs2 = sc.doc.Objects.FindByLayer('Legend')
    base_point = Rhino.Geometry.Point3d(0, 0, 0)
    idef_name2 = 'Blokk_Legend'

    geometry2 = [item.Geometry for item in objrefs2]
    attributes2 = [item.Attributes for item in objrefs2]

    idef_index = sc.doc.InstanceDefinitions.Add(idef_name2, "", base_point, geometry2, attributes2)
    rs.DeleteObjects(objrefs2)

def add_layers(layers_dict):
    for layer_name, color in layers_dict.items():
        rs.AddLayer(layer_name, color)

def symboler(tunnelStart, tunnelStop):
    for i in range(((tunnelStop-tunnelStart)//200)+1):
        rs.CurrentLayer('Berg_skadereg')
        rs.AddText('F1', (-100, i * 200),2)
        rs.AddText('F3', (-95, i * 200),2)
        rs.AddText('F7', (-90, i * 200),2)

        rs.CurrentLayer('Fremkommelighet')
        rs.AddText('1/X', (-100, -5+(i * 200)),2)
        rs.AddText('X', (-90, -5+(i * 200)),2)

def excel_import(excel):

    if os.path.isfile(excel):
        df_obs = pd.read_excel(excel, header=0)

        print(df_obs['Kommentar'])
        T_scale = 200
        vl_max = 110 / 500 * T_scale

    for index, row in df_obs.iterrows():

            try:
                pel = float(row['Pel'])
                pos = row['Pos']

                try:
                    if pos > -100 or pos < 100:
                        pos = float(pos) * vl_max / 100
                except:
                    print('Pos er ikke innenfor kartleggingsskjema, symbol lagt til origo')
                    pos = 0.0
                    pel = 0.0


                rs.AddText(row['Symbol'],(pos,pel), height=1.5, justification=2)
                rs.AddText(row['Kommentar'],(pos,pel-2), height=1.5, justification=2)

            except:
                print("Error:", row['Pel'], row['Kommentar'])


if __name__=="__main__":
    if rs.IsBlock('Blokk_Overskrift'):
        rs.DeleteBlock('Blokk_Overskrift')
    if rs.IsBlock('Blokk_Legend'):
        rs.DeleteBlock('Blokk_Legend')

    origin = Rhino.Geometry.Point3d.Origin
    normal = Rhino.Geometry.Vector3d.ZAxis

    direction = Rhino.Geometry.Vector3d(15, 255, 0)
    radians = Rhino.RhinoMath.ToRadians(0)

    t = Rhino.Geometry.Transform.Translation(direction)
    s = Rhino.Geometry.Transform.Scale(origin, 1.0)
    r = Rhino.Geometry.Transform.Rotation(radians, normal, origin)
    xform = t * s * r

    direction = Rhino.Geometry.Vector3d(135, 245, 0)
    t2 = Rhino.Geometry.Transform.Translation(direction)
    r2 = Rhino.Geometry.Transform.Rotation(radians, normal, origin)
    s2 = Rhino.Geometry.Transform.Scale(origin, 0.8)
    xform2 = t2 * s2 * r2

    tunnelNavn = rs.GetString('Tunnelnavn',)
    tunnelStart = rs.GetInteger('Tunnel start')
    tunnelStop = rs.GetInteger('Tunnel slutt')
    # excel = rs.GetString('PATH til Excel-fil')

    if tunnelStart >= tunnelStop:
        print("Startnummer må være mindre enn sluttnummer")


    add_layers(dwgLag())
    CreateOverskift(tunnelNavn)
    CreateLegend()
    lagTunnellinjer(tunnelStart, tunnelStop)
    symboler(tunnelStart, tunnelStop)
    # excel_import(excel)


    idefOverskrift = sc.doc.InstanceDefinitions.Find('Blokk_Overskrift')
    rs.CurrentLayer('Legend')
    idefLegend = sc.doc.InstanceDefinitions.Find('Blokk_Legend')


    layout_interval = 200
    layout_count = int(((tunnelStop - tunnelStart) / layout_interval)+1)

    for i in range(layout_count):
        # Sett inn ny Layout med riktig navn. 
        layout_y_start = tunnelStart+ (i * layout_interval)
        layout_y_end = layout_y_start + layout_interval
        layout_name = f"{int(layout_y_start)}-{int(layout_y_end)}"
        newLayout = rs.AddLayout(layout_name, (210,297))
        
        # Sett inn blokker med overskrift og tegneforklaring. 
        sc.doc.Objects.AddInstanceObject(idefOverskrift.Index, xform)
        sc.doc.Objects.AddInstanceObject(idefLegend.Index, xform2)

        # Sett inn detaljvindu og zoom til riktig plass. 
        detailGUID = rs.AddDetail(newLayout, (15, 35), (130, 255), 'Plan')

        pMin = rs.CreatePoint((0, layout_y_start+100, 0))

        detail = rs.coercerhinoobject(detailGUID)
        detailGeometry = detail.DetailGeometry
        detailViewPort = detail.Viewport
        detailView = detailViewPort.ParentView
        detailViewPort.SetCameraTarget(pMin, True)
        detailViewPort.Magnify(50,False)

        unit1 = Rhino.UnitSystem(4)
        unit2 = Rhino.UnitSystem(2)

        detailGeometry.SetScale(1,unit1,1,unit2)
        detail.CommitChanges()
        detail.CommitViewportChanges()

        sc.doc.Views.Redraw()

    rs.CurrentLayer("Default")
    view = rs.CurrentView('Top')
    rs.IsViewMaximized(view)
    rs.ZoomExtents()

