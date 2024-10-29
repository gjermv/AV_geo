#! python3
# -*- coding: UTF-8 -*-
# r: pandas, openpyxl


import Rhino
import rhinoscriptsyntax as rs
import scriptcontext as sc
import math
import System

import os

def tegneforklaring():
    tf = dict()
    tf['Berg (F)'] = 'X_Bold'
    tf['F1'] = 'Nedfall d<0.3m'
    tf['F2'] = 'Nedfall d>0.3m'
    tf['F3'] = 'Avløste blokker'
    tf['F4'] = 'Bom'
    tf['F5'] = 'Avskalling og bergslag'
    tf['F6'] = 'Utpressing'
    tf['F7'] = 'Vanninntrengning'
    tf['F8'] = 'Iskjøving'
    tf['Sprøytebetong (S)'] = 'X_Bold'
    tf['S1'] = 'Nedfall'
    tf['S2'] = 'Riss'
    tf['S3'] = 'Sprekker'
    tf['S4'] = 'Bom'
    tf['S5'] = 'Avskalling'
    tf['S6'] = 'Utpressing'
    tf['S7'] = 'Vanninntrengning'
    tf['S8'] = 'Iskjøving'
    tf['S9'] = 'Nedbrytning (vannkjemi, bakterier)'
    tf['Bolter til bergsikring (B)'] = 'X_Bold'
    tf['B1A-E'] = 'Korrosjon, Rustgrad A-E'
    tf['B2'] = 'Vrakbolt'
    tf['B3'] = 'Utpressing'
    tf['B4'] = 'Deformasjon'
    tf['Øvrige skader/mangler (M)'] = 'X_Bold'
    tf['M1'] = 'Manglende / ikke utført bergsikring'
    tf['M2'] = 'Mangler ved utført bergsikring'
    tf['M3'] = 'Manglende vedlikeholdsrensk'
    tf['M4'] = 'Skader på vann-/ frostsikringshvelv'
    tf['Framkommelighet'] = 'X_Bold'
    tf['X'] = ' Stengt'
    tf['I/X'] = 'Trangt'
    tf[' '] = ' '
    tf['Anbefalte tiltak'] = 'X_Bold'
    tf['Re'] = 'Rensk'
    tf['B'] = 'Bolt'
    tf['FB'] = 'Fjellbånd'
    tf['SFR'] = 'Sprøytebetong'
    tf['V'] = 'Vann/frostsikring'
    tf['A'] = 'Annet'
    
    return tf

def dwgLag():
    lag = {
        'Legend': (0, 0, 0),
        'Berg_skadereg': (255, 0, 0),
        'Vann-og frostsikring_skadereg': (0, 0, 255),
        'Kommentar': (0, 0, 0),
        'Forsagere': (255, 0, 0),
        'Fremkommelighet': (0, 0, 0),
        'Manglende utført bergsikring': (0, 0, 0),
        'Bolt': (0, 0, 0),
        'Plasser manuelt': (0,255,0),
    }

    return lag

def lagTunnellinjer(start, stop):
    rs.AddLayer("Kartleggingsskjema")
    rs.CurrentLayer("Kartleggingsskjema")

    T_scale = 200
    vl_max = 110/ 500 *T_scale 
    vl_mid = 50/ 500 *T_scale 
    hPos = [-vl_max, -vl_mid, 0, vl_mid, vl_max]  
    
    for h in hPos:
        rs.AddLine((h, start),(h, stop))

    rs.AddLine((-vl_max, start), (vl_max, start))
    rs.AddLine((-vl_max, stop), (vl_max, stop))
    rs.AddText(str(start),(-vl_max-5, start), 2, font=None, font_style=0, justification=131072+4)
    rs.AddText(str(stop),(-vl_max-5, stop), 2, font=None, font_style=0, justification=131072+4)
    
    vLineStart = int(((start // 20) + 1) * 20)
    for v in range(vLineStart, stop, 20):
        rs.AddLine((-vl_max,v),(vl_max,v))
        rs.AddText(str(v),(-vl_max-5,v), 2, font=None, font_style=0, justification=131072+4)


def CreateOverskift(tunnelNavn):

    tabellBredde = 180
    tabellHoyde = 10
    linjeA = 80
    linjeB = 110
    textBuffer = 1
    

    rs.CurrentLayer('Default')
    rs.AddText('Vedlegg 1 - Kartleggingsskjema', (textBuffer, 14.5), 3)

    rs.AddText('Tunnel (og løp):', (textBuffer, 8.5), 1.5)
    rs.AddText('Dato:', (linjeA+textBuffer, 8.5), 1.5)
    rs.AddText('Registert av:', (linjeB+textBuffer, 8.5), 1.5)
    
    rs.AddText(tunnelNavn, (textBuffer, 5), 3)

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

    sc.doc.InstanceDefinitions.Add(idef_name, "", base_point, geometry, attributes)
    rs.DeleteObjects(objrefs)


def CreateLegend():
    rs.CurrentLayer('Legend')
    legend = tegneforklaring()
    radHoyde = 4
    tabellBredde = 60
    tabellHoyde = (len(legend)) * radHoyde
    skriftStr = 2
    tabA = 2
    tabB = 12

    for i, key in enumerate(legend):
        rs.AddLine((0, 0 - (i * radHoyde)), (tabellBredde, 0 - (i * radHoyde)))
        if legend[key] == 'X_Bold':
            rs.AddText(key, (tabA, -0.85 - (i * radHoyde)), height=skriftStr, font=None, font_style=1)
        else:
            rs.AddText(key, (tabA, -0.85 - (i * radHoyde)), skriftStr)
            rs.AddText(legend[key], (tabB, -0.85 - (i * radHoyde)), skriftStr)

    rs.AddLine((0, -tabellHoyde), (tabellBredde, -tabellHoyde))

    rs.AddLine((0, 0), (tabellBredde, 0))
    rs.AddLine((0, 0), (0, -tabellHoyde))
    rs.AddLine((tabellBredde, 0), (tabellBredde, -tabellHoyde))

    rs.AddText('Kommentarer:', (tabA, -167.5), 2, font_style=1)
    rs.AddLine((0, -166), (tabellBredde, -166))
    rs.AddLine((0, -171), (tabellBredde, -171))
    rs.AddLine((0, -171-49), (tabellBredde, -171-49))
    rs.AddLine((0, -166), (0, -171-49))
    rs.AddLine((tabellBredde, -166), (tabellBredde, -171-49))

    objrefs2 = sc.doc.Objects.FindByLayer('Legend')
    base_point = Rhino.Geometry.Point3d(0, 0, 0)
    idef_name2 = 'Blokk_Legend'

    geometry2 = [item.Geometry for item in objrefs2]
    attributes2 = [item.Attributes for item in objrefs2]

    sc.doc.InstanceDefinitions.Add(idef_name2, "", base_point, geometry2, attributes2)
    rs.DeleteObjects(objrefs2)


def opprettLag(layers_dict):
    for layer_name, color in layers_dict.items():
        rs.AddLayer(layer_name, color)


def symboler(tunnelStart, tunnelStop):
    berg = ["F1", "F2", "F4", "F5", "F6", "S1", "S4", "S5", "S6"]
    vann = ["S7", "S3", "F7", "F8", "S2", "S8", "S9", "F9"]
    bolt = ['B1', 'B2', 'B3', 'B4', 'B1A', 'B1B', 'B1C', 'B1D', 'B1E']
    mangler = ['M1', 'M2', 'M3', 'M4', 'M5']
    fremkom = ["I/X", "X", "L", "D"]


    rs.CurrentLayer('Berg_skadereg')
    for j in range(len(berg)):
        rs.AddText(berg[j], (-88-(j*9), tunnelStart),3)

    rs.CurrentLayer('Vann-og frostsikring_skadereg')
    for h in range(len(vann)):
        rs.AddText((vann[h]), (-88-(h*9), tunnelStart-7), 3)

    rs.CurrentLayer('Bolt')
    for k in range(len(bolt)):
        rs.AddText(bolt[k], (-88-(k*14), tunnelStart-14), 3)

    rs.CurrentLayer('Manglende utført bergsikring')
    for p in range(len(mangler)):
        rs.AddText(mangler[p], (-88-(p*9.2), tunnelStart-21), 3)

    rs.CurrentLayer('Fremkommelighet')
    for t in range(len(fremkom)):
        rs.AddText(fremkom[t], (-88-(t*8), tunnelStart+-28),3)

    rs.CurrentLayer('Forsagere')
    rs.AddText('!!!', (-88, tunnelStart-35),3)
    rs.AddCircle((-85.25, tunnelStart+(-37.2)),3.7)

def xform2DLayout(x, y): 
    origin = Rhino.Geometry.Point3d.Origin
    normal = Rhino.Geometry.Vector3d.ZAxis

    direction = Rhino.Geometry.Vector3d(x, y, 0)
    radians = Rhino.RhinoMath.ToRadians(0)

    t = Rhino.Geometry.Transform.Translation(direction)
    s = Rhino.Geometry.Transform.Scale(origin, 1.0)
    r = Rhino.Geometry.Transform.Rotation(radians, normal, origin)
    xform = t * s * r

    return xform

def slettEksisterendeBlokker():
    if rs.IsBlock('Blokk_Overskrift'):
        rs.DeleteBlock('Blokk_Overskrift')
    if rs.IsBlock('Blokk_Legend'):
        rs.DeleteBlock('Blokk_Legend')

def hentTunnelinfo():
    tunnelNavn = rs.GetString('Tunnelnavn')
    tunnelStart = rs.GetInteger('Tunnel start')
    tunnelStop = rs.GetInteger('Tunnel slutt')

    if tunnelStart >= tunnelStop:
        tunnelStart, tunnelStop = tunnelStop, tunnelStart

    return [tunnelNavn, tunnelStart, tunnelStop]

if __name__=="__main__":
    slettEksisterendeBlokker()

    tunnelNavn, tunnelStart, tunnelStop = hentTunnelinfo()

    opprettLag(dwgLag())
    CreateOverskift(tunnelNavn)
    CreateLegend()
    lagTunnellinjer(tunnelStart, tunnelStop)
    symboler(tunnelStart, tunnelStop)


    idefOverskrift = sc.doc.InstanceDefinitions.Find('Blokk_Overskrift')
    idefLegend = sc.doc.InstanceDefinitions.Find('Blokk_Legend')


    layout_interval = 200
    layout_count = int(((tunnelStop - tunnelStart) / layout_interval)+1)

    for i in range(layout_count):
        # Sett inn ny Layout med riktig navn. 
        layout_y_start = tunnelStart+ (i * layout_interval)
        layout_y_end = layout_y_start + layout_interval
        layout_name = f"{int(layout_y_start)}-{int(layout_y_end)}"
        newLayout = rs.AddLayout(layout_name, (210,297))
        
        # Sett inn blokker med overskrift og tegneforklaring
        rs.CurrentLayer('Legend')
        sc.doc.Objects.AddInstanceObject(idefLegend.Index, xform2DLayout(135, 255))
        rs.CurrentLayer('Default')
        sc.doc.Objects.AddInstanceObject(idefOverskrift.Index, xform2DLayout(15,263))
        
        # Sett inn detaljvindu og zoom til riktig plass. 
        detailGUID = rs.AddDetail(newLayout, (15, 35), (130, 255), 'Plan')

        pMin = rs.CreatePoint((-5, layout_y_start + 100, 0))

        detail = rs.coercerhinoobject(detailGUID)
        detailGeometry = detail.DetailGeometry
        detailViewPort = detail.Viewport
        detailView = detailViewPort.ParentView
        detailViewPort.SetCameraTarget(pMin, True)
        detailViewPort.Magnify(50, False)

        unit1 = Rhino.UnitSystem(4)
        unit2 = Rhino.UnitSystem(2)

        detailGeometry.SetScale(0.96, unit1, 1, unit2)
        detail.CommitChanges()
        detail.CommitViewportChanges()

        sc.doc.Views.Redraw()

    rs.CurrentLayer("Default")
    view = rs.CurrentView('Top')
    rs.IsViewMaximized(view)
    rs.ZoomExtents()

