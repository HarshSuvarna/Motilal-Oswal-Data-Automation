import pandas as pd
import panel as pn
import hvplot.pandas
import holoviews as hv
from IPython import get_ipython
import pandas as pd
import xlsxwriter
import os
import glob
import pandas as pd

path = os.getcwd()
files = glob.glob(os.path.join(path, '*.xlsx'))
csv_files = glob.glob(os.path.join(path, '*csv'))


hv.extension('bokeh')
pn.extension('tabulator', sizing_mode="stretch_width")

if files[0].split('.')[1] == 'xlsx':
    data = pd.read_excel(files[0])
    
else:
    data = pd.read_csv(files[0])
    


def environment():
    try:
        get_ipython()
        return "notebook"
    except:
        return "server"
environment()

idata = data.interactive()

text = pn.widgets.TextInput(
    name='Subtype Code',
    placeholder='Enter Code here')

pn.Row(text, height=100)

def filters(text):
    if text:
        if (text in idata['FASUBTYPECODE'].unique()):
            return idata[idata['FASUBTYPECODE']==text]['FASUBTYPECODE'].count(), idata[idata['FASUBTYPECODE']==text]
    

ipipeline = (filters(text))


if environment()=="server":
    theme="fast"
else:
    theme="simple"

itable = ipipeline[1].pipe(pn.widgets.Tabulator, pagination='remote', page_size=21, theme=theme)

template = pn.template.FastListTemplate(
    title='Broker Client Information', 
    sidebar=[text, filters(text)[0]],
    main=[itable.panel()],
    accent_base_color="#88d8b0",
    header_background="#deac1b",
)
template.show()
