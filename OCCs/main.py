from tkinter import *
import tkinter as tk
import pandas as pd
import glob2
from tkinter import messagebox, filedialog, ttk
import re
import ifcopenshell.util
import ifcopenshell.util.element
import xlsxwriter
import ifcopenshell
from ifcopenshell import geom
geom.settings()
import ifcopenshell.geom
import os
from OCC.Core.gp import gp_Vec
from OCC.Core.Quantity import Quantity_Color, Quantity_TOC_RGB
from OCC.Core.Graphic3d import Graphic3d_ClipPlane

from OCC.Display.SimpleGui import init_display


pd.set_option('display.max_colwidth', None)
root = tk.Tk()
root.geometry("1200x750")
root.pack_propagate(False)
root.resizable(0, 0)

#Menu för rows eller columns
val = StringVar()
val.set("Columns")
dragmenu = OptionMenu(root, val, "Rows", "Columns" )
dragmenu.pack()
dragmenu.place(rely=0.85, relx=0.50)

#Rutan som excel texten är i
frame1 = tk.LabelFrame(root, text="Excel Output")
frame1.place(height = 350, width = 1200)

#Rutan för import av saker
file_frame = tk.LabelFrame(root, text="IFC")
file_frame.place(height=100, width=350, rely=0.60, relx=0)
file_frame1 = tk.LabelFrame(root, text="Excel")
file_frame1.place(height=100, width=350, rely=0.80, relx=0)

#Status om vad det finns för columns
columntext1 = tk.Label(root, text = "")
columntext1.place(rely=0.68, relx=0.47)
columntext2 = tk.Label(root, text = "")
columntext2.place(rely=0.74, relx=0.47)
#Skriva in sökord

#Knappar
button1 = tk.Button(file_frame1, text="Flera excel i mapp", command=lambda: file_dialog())
button1.place(rely=0.65, relx=0.50)

button2 = tk.Button(file_frame1, text="Sök", command=lambda: ladda_excel())
button2.place(rely=0.65, relx=0.90)

button2 = tk.Button(root, text="3d-view", command=lambda: dview())
button2.place(rely=0.5, relx=0.20)

button3 = tk.Button(file_frame, text="Välj IFC-fil för Excel", command=lambda: ladda_IFC())
button3.place(rely=0.65, relx=0.30)
button3 = tk.Button(file_frame, text="Skapa Excel av IFC", command=lambda: Skapa_excel())
button3.place(rely=0.65, relx=0.65)
button3 = tk.Button(root, text="Rensa Sökning", command=lambda: rensa_data())
button3.place(rely=0.50, relx=0)

#Inte Aktiv Än
label_file = ttk.Label(file_frame1, text="Ingen Xlsx Vald")
label_file.place(rely=0,relx=0)

label_file1 = ttk.Label(file_frame, text="Ingen IFC Vald")
label_file1.place(rely=0,relx=0)

#Display av excel
tv1 = ttk.Treeview(frame1)
#Så hela excel är i treeviewen
tv1.place(relheight=1,relwidth=1)
#Scroll ifall det är förmycket
treescrolly= tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)
treescrollx= tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)

treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")

def ladda_IFC():
    #Måste läsa så rutan resetats efter sökningen en till knapp ?
    #tv1.delete()
    filenames = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File", filetype=(("IFC files", "*.ifc"),("All Files", "*.*")))
    global Filnamn
    Filnamn = filenames
    filename_noIFC = filenames.removesuffix(".ifc")
    head, tail = os.path.split(filename_noIFC)
    label_file1["text"] = tail
def ladda_excelfil():
    #Måste läsa så rutan resetats efter sökningen en till knapp ?
    #tv1.delete()
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Välj en Fil", filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))

    dk = pd.read_excel(filename)
    rows = val.get()
    #Detta fungerar inte med Rows än get ut 0 1 2 som namn och inte raderna
    if rows == "Rows":
        dk = dk.T
        dk.drop(columns=dk.columns[0],
                axis=1,
                inplace=True)
        print(dk)
    ord = list(dk.columns)
    var = tk.StringVar()
    l = int(len(ord))
    k = 0
    #Loopar igenom excelfilen och sklapar en ruta för varje column/rad som headar och fixar ett sökfält
    for i in range(l):
        if len(ord[0]) > 0:
            globals()['var%s' % k] = tk.StringVar()
            columntext1 = tk.Label(root, text="{0}".format(ord[0+k]))
            columntext1.place(rely=(0.50 + (k*6/100)), relx=0.75)
            globals()['ent%s' % k] = tk.Entry(root, textvariable= globals()['var%s' % k])
            globals()['ent%s' % k].place(height=20, width=70, rely=(0.5 + (k*6/100)), relx=0.85)
            k += 1



def file_dialog():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Välj en File", filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    dk = pd.read_excel(filename)
    rows = val.get()
    #Rows är inte helt fixade
    if rows == "Rows":
        dk = dk.T
        dk = dk.set_index(1)
        k = 0
    else:
        k = 0
    ord = list(dk.columns)
    print(ord)
    var = tk.StringVar()
    l = int(len(ord))
    for i in range(l):
        globals()['var%s' % k] = tk.StringVar()
        columntext1 = tk.Label(root, text="{0}".format(ord[0+k]))
        columntext1.place(rely=(0.50 + (k*3/100)), relx=0.75)
        globals()['ent%s' % k] = tk.Entry(root, textvariable= globals()['var%s' % k])
        globals()['ent%s' % k].place(height=20, width=70, rely=(0.50 + (k*3/100)), relx=0.85)
        k += 1
    path = filedialog.askdirectory(title = 'Select directory with files to decrypt')
    if path == None or path == '':
        messagebox.showwarning('Error', 'No path selected, exiting...')
        return False
    path =  path + '/'
    label_file["text"]=path
def ladda_excel():

    path = os.getcwd()
    csv_files = glob2.glob(os.path.join(r"{0}".format(label_file["text"]), "*.xlsx"))
    k = 0
    i = 0
    for f in csv_files:
        #Läser in Excel filen
        df = pd.read_excel(f)
    rows = val.get()
    k=0
    ord = list(df.columns)
    l = int(len(ord))
    #Så rutan anpassas
    a = ""
    #Den klarar inte av int / digits värden från Excelet än.
    while i < l:
        if len(globals()['ent%s' % k].get()) < 1:
            globals()['vaUpper%s' % k] = globals()['ent%s' % k].get()
            globals()['vaUpper%s' % k] = re.compile(r'([+-]?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?|[^\W_])|.', re.DOTALL)
        else:
            globals()['va%s' % k] = globals()['ent%s' % k].get()
            globals()['vaUpper%s' % k] = globals()['va%s' % k].upper()
        k += 1
        i += 1
    for f in csv_files:
        #Läser in Excel filen och göra alla värden till ints, så date och nummer räknas med när den söker
        df = pd.read_excel(f, dtype=str)
        print(df)
        rows = val.get()
        if rows == "Rows":
            df = df.T
            tv1["column"] = list(df.columns)
            tv1["show"] = "headings"
            ord = list(df.T.columns)
            # Fixar så att "Sök headline" är överst
            for row in tv1["column"]:
                tv1.heading(row, text=row)
                ord = list(df.columns)
        else:
            #df.set_index("Start", inplace=True)
            tv1["column"] = list(df.columns)
            tv1["show"] = "headings"
            ord = list(df.columns)

            for column in tv1["column"]:
                tv1.heading(column, text=column)
                ord = list(df.columns)

        k = 0
        b = 1
        for i in range(len(ord)):
            #Fixar så tomma celler inte är tomma
            df.fillna(' ', inplace=True)
            HJ = df.loc[df["{}".format(ord[k])].isnull()]

            #Säger hur raderna ska sökas igenom ord[k] raden av den columnen tittar ifall str.cointains har med ordet som har skrivits in innan
            globals()['a%s' % k] = df["{}".format(ord[k])].str.upper().str.contains(globals()['vaUpper%s' % k])
            i +=1
            b += 1
            k += 1
        # Fixar så att "vart/vad/problem är överst

        k = 0
        #Tittar igenom alla de olika raderna och lägger till de i en array liknande tupel som blir c
        for i in range(len(ord)):
            if k == 0:
                a = globals()['a%s' % k]
                k += 1
                i += 1
            else:
                b = globals()['a%s' % k]
                k += 1
                i += 1
                #Adderar b till varje rad med & och inte + för att få unika b till varje
                a = a & b
            c = (df[a])
        '''    
        if len(ord) <= 1:
            a = (df[df["{0}".format(ord[0])].str.contains(globals()['vaUpper%s' % k])])
        else:

            a = (df[df["{0}".format(ord[0])].str.contains(globals()['vaUpper%s' % 1]) + df["{0}".format(ord[1])].str.contains(globals()['vaUpper%s' % 2])+ df["{0}".format(ord[2])].str.contains(globals()['vaUpper%s' % 3])+ + df["{0}".format(ord[3])].str.contains(globals()['vaUpper%s' % 4])])
        '''
        title = (f.split("\\")[-1])
        strtitle = str(title)
        print(title)
        tv1.insert("", "end", values=strtitle)
        #Vad ska printas ut
        if len(a) > 0:
                #Removar så den bara tar med rader som har sögorden i sig
            df_rows = c.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)


def rensa_data():
        tv1.delete(*tv1.get_children())


def dview():
    filename = Filnamn
    ifc_file = ifcopenshell.open(filename)
    display, start_display, add_menu, add_function_to_menu = init_display()

    settings = ifcopenshell.geom.settings()
    settings.set(
        settings.USE_PYTHON_OPENCASCADE, True
    )  # tells ifcopenshell to use pythonocc

    # read the ifc file


    # the clip plane

    # clip plane number one, by default xOy
    clip_plane_1 = Graphic3d_ClipPlane()

    # set hatch on
    clip_plane_1.SetCapping(True)
    clip_plane_1.SetCappingHatch(True)

    # off by default, user will have to enable it
    clip_plane_1.SetOn(False)

    # set clip plane color
    aMat = clip_plane_1.CappingMaterial()
    aColor = Quantity_Color(0.5, 0.6, 0.7, Quantity_TOC_RGB)
    aMat.SetAmbientColor(aColor)
    aMat.SetDiffuseColor(aColor)
    clip_plane_1.SetCappingMaterial(aMat)

    # and display each subshape
    products = ifc_file.by_type("IfcProduct")  # traverse all IfcProducts
    nb_of_products = len(products)
    for i, product in enumerate(products):
        if (
                product.Representation is not None
        ):  # some IfcProducts don't have any 3d representation
            try:
                pdct_shape = ifcopenshell.geom.create_shape(settings, inst=product)
                r, g, b, a = pdct_shape.styles[0]  # the shape color
                color = Quantity_Color(abs(r), abs(g), abs(b), Quantity_TOC_RGB)
                # speed up rendering, don't update rendering for each shape
                # only update all 50 shapes
                to_update = i % 50 == 0
                new_ais_shp = display.DisplayShape(
                    pdct_shape.geometry,
                    #Ändra color = color ifall man vill ha med texturer till 3dmodellen men detta öka kravet för dator.
                    color=color,
                    transparency=abs(1 - a),
                    update=to_update,
                )[0]
                new_ais_shp.AddClipPlane(clip_plane_1)
            except RuntimeError:
                print("Failed to process shape geometry")

    def animate_translate_clip_plane(event=None):
        clip_plane_1.SetOn(True)
        plane_definition = clip_plane_1.ToPlane()  # it's a gp_Pln
        h = 0.01
        for _ in range(1000):
            plane_definition.Translate(gp_Vec(0.0, 0.0, h))
            clip_plane_1.SetEquation(plane_definition)
            display.Context.UpdateCurrentViewer()

    if __name__ == "__main__":
        add_menu("IFC clip plane")
        add_function_to_menu("IFC clip plane", animate_translate_clip_plane)
        display.FitAll()
        start_display()

def Skapa_excel():
    filename = Filnamn
    ifc_file = ifcopenshell.open(filename)
    filename_noIFC = filename.removesuffix(".ifc")
    head, tail = os.path.split(filename_noIFC)
    path = filedialog.askdirectory(title='Select directory with files to decrypt')
    workbook = xlsxwriter.Workbook("{}/{}.xlsx".format(path, tail))
    worksheet = workbook.add_worksheet("Sheet 1")
    worksheet.write(0, 0, 'Område')
    worksheet.write(0, 1, 'Familjer Bärande')
    worksheet.write(0, 2, 'Familjer EjBärande')
    worksheet.write(0, 3, 'UE för Arbete')
    worksheet.write(0, 4, 'Material')
    worksheet.write(0, 5, 'Mängd Material')
    worksheet.write(0, 6, 'Våningar')
    worksheet.write(0, 7, 'Area')
    worksheet.write(0, 8, 'Volym')
    worksheet.write(0, 9, 'Kostnad')
    worksheet.write(0, 10, 'Kommentarer')
    print(path)
    print("{}/{}.xlsx".format(path, tail))

    # wb.save("{0}/testar.xlsx".format(path))
    # Får ut alla olika familjer i IFC filen
    element = ifc_file.by_type('IfcElement')
    Loadproducts = []
    NoneLoadproducts = []

    def print_propertiesLoad(properties):
        for name, value in properties.items():
            # Reference är vilken familj den tar av kan få ut id, ifctyp har också, checkar om elementet redan finns i listan annars appendas det till
            if name == "Reference":
                value = "{}".format(ps) + " -> " + value
                Loadproducts.append(value) if value not in Loadproducts else Loadproducts
        # Kanske ha med avd för familj det tillhör innan

    # Så man bara får alla loadbearings delar
    def print_propertiesNoneLoad(properties):
        for name, value in properties.items():
            # Reference är vilken familj den tar av kan få ut id, ifctyp har också, checkar om elementet redan finns i listan annars appendas det till
            if name == "Reference":
                value = "{}".format(ps) + " -> " + value
                NoneLoadproducts.append(value) if value not in NoneLoadproducts else NoneLoadproducts

    from ifcopenshell.util.element import get_psets
    for elements in element:
        for ps, p in get_psets(elements).items():
            for name, value in p.items():
                if name == "LoadBearing":
                    if value == True:
                        print_propertiesLoad(p)
                    else:
                        print_propertiesNoneLoad(p)
#Är här som mängdningen ska komma in.
    products = ifc_file.by_type('IfcElement')
    material_list = []
#If "A" pga om jag behöver göra olika sökningar för IFC och IFC2x3
    if ifc_file.schema != "IFC4":
        for product in products:
            if product.HasAssociations:
                for i in product.HasAssociations:
                    if i.is_a('IfcRelAssociatesMaterial'):

                        if i.RelatingMaterial.is_a('IfcMaterialSelect'):
                            material_list.append(
                                materials.Material.Name) if materials.Material.Name not in material_list else material_list
#IfcMaterialList har inte Material så behöver bara materials.Name
                        if i.RelatingMaterial.is_a('IfcMaterialList'):
                            for materials in i.RelatingMaterial.Materials:
                                material_list.append(
                                    materials.Name) if materials.Name not in material_list else material_list

                        if i.RelatingMaterial.is_a('IfcMaterialLayerSetUsage'):
                            for materials in i.RelatingMaterial.ForLayerSet.MaterialLayers:
                                material_list.append(
                                    materials.Material.Name) if materials.Material.Name not in material_list else material_list

    else:
        material = ifc_file.by_type('IfcMaterial')

        for materials in material:
            string_mat = str(materials)
            only_mat = string_mat.partition("IfcMaterial")[-1]
            material_list.append(only_mat) if only_mat not in material_list else material_list

    # Bara för att visa vad som fungerar x ska bli inputet till excelfilen
    print("LoadBearing Elements:")
    l = 1
    nl = 1
    mt = 1
    for m in (material_list):
        worksheet.write(mt, 4, m)
        mt += 1
    for x in sorted(Loadproducts):
        worksheet.write(l, 1, x)
        l += 1
    print("NoneLoadBearing Elements:")
    for y in sorted(NoneLoadproducts):
        worksheet.write(nl, 2, y)
        nl += 1
    workbook.close()


#Måste göra en till def för Area och Volym area för golv + vägga volym för resten tror jag.

root.mainloop()