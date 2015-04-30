# -*- coding: utf-8 -*-
"""
This script takes a Tableau workbook file (*twb) as input and returns an excel
(*xlsx) file with up to five sheets: Fields, Parameters, Dashboards, Actions 
and Misc. The returned Excel file can be used as a jumping off point for a data
dictionary. 


Created on Mon Apr 27 19:38:46 2015


Known Issues: 
    - Only reports fields used on worksheets


@author: 51nc17y
"""
# ONLY INPUT REQUIRED:
#*****************************************************************************
#file_twb = r"C:\Temp\Coffee.twb"    
#*****************************************************************************

#TEST CASES:
#file_twb = r"C:\Temp\Superstore.twb"
#file_twb = r"C:\Temp\Coffee.twb" 
#file_twb = r"C:\Temp\Shelter.twb"
#file_twb = r"C:\Temp\Soccer.twb"
file_twb = r"C:\Temp\Gazprom.twb"

import os
import shutil
from pandas import DataFrame, ExcelWriter

try:
    import cElementTree as ET
except ImportError:
    try:
    # Python 2.5 need to import a different module
        import xml.etree.cElementTree as ET
    except ImportError:
        exit_err("Failed to import cElementTree from any known place")

fileName, fileExt = fileName, fileExt = os.path.splitext(os.path.basename(file_twb))

#Creating XML file from TWB data 
file_xml = os.path.dirname(file_twb) +'\\'+fileName+'.xml'
shutil.copy2(file_twb, file_xml)

#Create tree structure from XML file
tree = ET.parse(file_xml)
root = tree.getroot()

#DASHBOARDS
if not root.find('dashboards'):
    print "WARNING: NO DASHBOARDS FOUND"
else: 
    #Initialize Dashboard Lists
    d_names = []
    d_dependency_datatypes = []
    d_dependency_roles = []
    d_dependency_names = []
    d_dependency_types = []
    d_datasource_captions = []
    d_datasource_names = []
    d_datasource_dependencies = []
    d_minheights = []
    d_minwidths = []
    d_maxheights = []
    d_maxwidths = []
    d_ids = []
    d_paths = []
    d_sites = []    
    
    #Extracting Dashboard Data
    for dashboard in range(0, len(root.find('dashboards'))):
        dependencies = root.find('dashboards')[dashboard].findall('.//column')
        if len(dependencies)==0:
            print ("WARNING: Dependencies are empty in Dashboard #"+str(dashboard+1))
            
            d_names.append(root.find('dashboards')[dashboard].attrib['name'])      
            #d_datasource_dependencies.append(root.find('dashboards')[dashboard][4].attrib['datasource'])
            #d_dependency_datatypes.append(dependency.attrib['datatype'])
            d_dependency_roles.append('N/A')
            d_dependency_names.append('N/A')
            d_dependency_types.append('N/A')
            
            try:       
                if root.find('dashboards')[dashboard].find('repository-location').attrib:
                    if 'id' in root.find('dashboards')[dashboard].find('repository-location').attrib:
                        d_ids.append(root.find('dashboards')[dashboard].find('repository-location').attrib['id'])
                    else: 
                        d_ids.append('N/A')
                    if 'path' in root.find('dashboards')[dashboard].find('repository-location').attrib:
                        d_paths.append(root.find('dashboards')[dashboard].find('repository-location').attrib['path'])
                    else: 
                        d_paths.append('N/A')
                    if 'site' in root.find('dashboards')[dashboard].find('repository-location').attrib:
                        d_sites.append(root.find('dashboards')[dashboard].find('repository-location').attrib['site'])
                    else: 
                        d_sites.append('N/A')
                else: 
                    d_sites.append('N/A')
                    d_paths.append('N/A')
                    d_ids.append('N/A')
            except AttributeError:
                print "WARNING: No repository location in Dashboard #" + str(dashboard)
                
            if 'maxheight' in root.find('dashboards')[dashboard].find('size').attrib:
                d_maxheights.append(root.find('dashboards')[dashboard].find('size').attrib['maxheight'])
            else:
                d_maxheights.append('N/A')
            if 'minheight' in root.find('dashboards')[dashboard].find('size').attrib:
                d_minheights.append(root.find('dashboards')[dashboard].find('size').attrib['minheight'])
            else:
                d_minheights.append('N/A')
            if 'maxwidth' in root.find('dashboards')[dashboard].find('size').attrib:
                d_maxwidths.append(root.find('dashboards')[dashboard].find('size').attrib['maxwidth'])
            else:
                d_maxwidths.append('N/A')
            if 'minwidth' in root.find('dashboards')[dashboard].find('size').attrib:
                d_minwidths.append(root.find('dashboards')[dashboard].find('size').attrib['minwidth'])
            else:
                d_minwidths.append('N/A')
        else:
            for dependency in dependencies:
                d_names.append(root.find('dashboards')[dashboard].attrib['name'])
                d_datasource_dependencies.append(root.find('dashboards')[dashboard][4].attrib['datasource'])
                d_dependency_datatypes.append(dependency.attrib['datatype'])
                d_dependency_roles.append(dependency.attrib['role'])
                d_dependency_names.append(dependency.attrib['name'])
                d_dependency_types.append(dependency.attrib['type'])
        
                if 'caption' in root.find('dashboards')[dashboard].find('datasources/datasource').attrib:
                    d_datasource_captions.append(root.find('dashboards')[dashboard].find('datasources/datasource').attrib['caption'])
                else:
                    d_datasource_captions.append('N/A')
                if 'name' in root.find('dashboards')[dashboard].find('datasources/datasource').attrib:
                    d_datasource_names.append(root.find('dashboards')[dashboard].find('datasources/datasource').attrib['name'])
                else:
                    d_datasource_names.append('N/A')
                    
                if 'maxheight' in root.find('dashboards')[dashboard].find('size').attrib:
                    d_maxheights.append(root.find('dashboards')[dashboard].find('size').attrib['maxheight'])
                else:
                    d_maxheights.append('N/A')
                if 'minheight' in root.find('dashboards')[dashboard].find('size').attrib:
                    d_minheights.append(root.find('dashboards')[dashboard].find('size').attrib['minheight'])
                else:
                    d_minheights.append('N/A')
                if 'maxwidth' in root.find('dashboards')[dashboard].find('size').attrib:
                    d_maxwidths.append(root.find('dashboards')[dashboard].find('size').attrib['maxwidth'])
                else:
                    d_maxwidths.append('N/A')
                if 'minwidth' in root.find('dashboards')[dashboard].find('size').attrib:
                    d_minwidths.append(root.find('dashboards')[dashboard].find('size').attrib['minwidth'])
                else:
                    d_minwidths.append('N/A')
                
            try:       
                if root.find('dashboards')[dashboard].find('repository-location').attrib:
                    if 'id' in root.find('dashboards')[dashboard].find('repository-location').attrib:
                        d_ids.append(root.find('dashboards')[dashboard].find('repository-location').attrib['id'])
                    else: 
                        d_ids.append('N/A')
                    if 'path' in root.find('dashboards')[dashboard].find('repository-location').attrib:
                        d_paths.append(root.find('dashboards')[dashboard].find('repository-location').attrib['path'])
                    else: 
                        d_paths.append('N/A')
                    if 'site' in root.find('dashboards')[dashboard].find('repository-location').attrib:
                        d_sites.append(root.find('dashboards')[dashboard].find('repository-location').attrib['site'])
                    else: 
                        d_sites.append('N/A')
                else: 
                    d_sites.append('N/A')
                    d_paths.append('N/A')
                    d_ids.append('N/A')
            except AttributeError:
                print "WARNING: No repository location in Dashboard #" + str(dashboard)
                
                    
    print str(len(root.findall('dashboards'))) + " Dashboards found"     
    
#ACTIONS
if not root.find('actions'):
    print "WARNING: NO ACTIONS FOUND"
else:  
    #Initialize Action Lists
    a_captions = []
    a_names = []
    a_activation_types = []
    a_activation_auto_clears = []
    a_source_dashboards = []
    a_source_worksheets = []
    a_source_types = []
    a_cmds = []
    a_cmd_param_names = []
    a_cmd_param_values = []
    
    #Extracting Action Data
    for action in range(0, len(root.find('actions'))):
        params = root.find('actions')[action].findall('.//param')
        for param in params:
            a_captions.append(root.find('actions')[action].attrib['caption'])
            a_names.append(root.find('actions')[action].attrib['name'])
            a_source_types.append(root.find('actions')[action].find('source').attrib['type'])
            a_cmd_param_names.append(param.attrib['name'])
            a_cmd_param_values.append(param.attrib['value'])
    
            if 'auto-clear' in root.find('actions')[action][0].attrib:
                a_activation_auto_clears.append(root.find('actions')[action][0].attrib['auto-clear'])
            else:
                a_activation_auto_clears.append('N/A')
            if 'worksheet' in root.find('actions')[action].attrib:
                a_source_worksheets.append(root.find('actions')[action].find('source').attrib['worksheet'])
            else:
                a_source_worksheets.append('N/A')
            if 'type' in root.find('actions')[action].attrib:
                a_activation_types.append(root.find('actions')[action].find('activation').attrib['type'])
            else:
                a_activation_types.append('N/A')
            if 'dashboard' in root.find('actions')[action].find('source').attrib:
                a_source_dashboards.append(root.find('actions')[action].find('source').attrib['dashboard'])
            else:
                a_source_dashboards.append('N/A')
                
    print str(len(root.find('actions'))) + " Actions found"

#PARAMETERS
if not root.find('datasources')[0].findall('column'):
    print "WARNING: NO PARAMETERS FOUND"
else:
    #Initialize Parameter Lists
    p_captions = []
    p_names = []
    p_formats = []
    p_datatypes = []
    p_domaintypes = []
    p_types = []
    p_roles = []
    p_values = []
    p_granularities = []
    p_mins = []
    p_maxs = []
    p_values = []
    p_formulas = []
    p_keyes = []
    p_key_values = []
    p_aggregation = []
    p_semantic_role = []
    
    
    params = root.find('datasources')[0].findall('column')
    
    for param in params:
        p_names.append(param.attrib['name'])
        p_datatypes.append(param.attrib['datatype']) 
        p_types.append(param.attrib['type'])
        p_roles.append(param.attrib['role'])
        
        if 'aggregation' in param.attrib:
            p_aggregation.append(param.attrib['aggregation'])        
        else:
            p_aggregation.append('N/A')
        if 'semantic-role' in param.attrib:
            p_semantic_role.append(param.attrib['semantic-role'])        
        else:
            p_semantic_role.append('N/A')
        if 'value' in param.attrib:
            p_values.append(param.attrib['value'])        
        else:
            p_values.append('N/A')
        if 'param-domain-type' in param.attrib:
            p_domaintypes.append(param.attrib['param-domain-type'])
        else:
            p_domaintypes.append('N/A')
        if 'caption' in param.attrib:
            p_captions.append(param.attrib['caption'])
        else:
            p_captions.append('N/A')
        if 'default-format' in param.attrib:
            p_formats.append(param.attrib['default-format'])
        else:
            p_formats.append('N/A')
        
        
            
        if param.find('calculation'):
            if 'formula' in param.find('calculation').attrib:
                p_formulas.append(param.find('calculation').attrib['formula'])
            else:
                p_formulas.append('N/A')
        else:
            p_formulas.append('N/A')
        
        if param.find('range'):
            if 'granularity' in param.find('range').attrib:
                p_granularities.append(param.find('range').attrib['granularity'])
            else:
                p_granularities.append('N/A')
            if 'min' in param.find('range').attrib:
                p_mins.append(param.find('range').attrib['min'])
            else:
                p_mins.append('N/A')
            if 'max' in param.find('range').attrib:
                p_maxs.append(param.find('range').attrib['max'])
            else:
                p_maxs.append('N/A')
        else:
             p_granularities.append('N/A')
             p_mins.append('N/A')
             p_maxs.append('N/A')
    
    print str(len(params)) + " Parameters found"

#Initialize Field Lists
f_datatypes = []
f_types = []
f_roles = []
f_names = []
f_captions = []
f_ws_names = []
f_formulas = []
f_default_formats = []

columns = root.findall('.//worksheet/table/view/datasource-dependencies/column')

for worksheet in range(0, len(root.find('worksheets'))):
    for column in columns:
        f_ws_names.append(root.find('worksheets')[worksheet].attrib['name'])
        f_datatypes.append(column.attrib['datatype'])
        f_types.append(column.attrib['type'])
        f_roles.append(column.attrib['role'])
        if 'caption' in column.attrib:
            f_captions.append(column.attrib['caption'])
        else:
            f_captions.append('N/A')
        if 'name' in column.attrib:
            f_names.append(column.attrib['name'])
        else:
            f_names.append('N/A')
        if 'formula' in column.attrib:
            f_formulas.append(column.attrib['formula'])
        else:
            f_formulas.append('N/A')
        if 'default-format' in column.attrib:
            f_default_formats.append(column.attrib['default-format'])
        else:
            f_default_formats.append('N/A')

if len(columns) == 0:
    print "WARNING: NO FIELDS FOUND"
else:
    print str(len(columns)) + " Fields found"

#Initialize Misc Lists          
m_names = ['Source Plattform: ', 'Tableau Version: ']
m_values = []

m_values.append(root.attrib['source-platform'])
m_values.append(root.attrib['version'])


#Getting the data in dataframes for export
d1 = {'Worksheet Name': f_ws_names,
      'Field Caption': f_captions,
      'Field Name': f_names,
      'Role': f_roles,
      'Datatype': f_datatypes,
      'Types': f_types,
      'Formula': f_formulas,
      'Default Format': f_default_formats}

d2 = {'Parameter Caption': p_captions,
      'Parameter Name': p_names,
      'Datatype': p_datatypes,
      'Domain Type': p_domaintypes,
      'Types': p_types,
      'Format': p_formats,
      'Role': p_roles,
      'Granularity': p_granularities,
      'Minumum': p_mins,
      'Default Value': p_values,
      'Maximum': p_maxs,
      'Formula': p_formulas,
      'Aggregation': p_aggregation,
      'Semantic Role': p_semantic_role}

if len(dependencies)==0:
    d3 = {'Dashboard Name': d_names,
          'ID': d_ids,
          'Path': d_paths, 
          'Site': d_sites,
          'Dashboard Min Height':d_minheights,
          'Dashboard Max Height': d_maxheights,
          'Dashboard Min Width': d_minwidths,
          'Dashboard Max Width': d_maxwidths}
else:
    d3 = {'Dashboard Name': d_names,
          'Dashboard Min Height':d_minheights,
          'Dashboard Max Height': d_maxheights,
          'Dashboard Min Width': d_minwidths,
          'Dashboard Max Width': d_maxwidths,
          'Datasource Caption' : d_datasource_captions,
          'Datasource Name': d_datasource_names,
          'Datasource Dependency' : d_datasource_dependencies,
          'Dependency Name': d_dependency_names,
          'Dependency Datatype': d_dependency_datatypes,
          'Dependency Role' : d_dependency_roles,
          'Dependency Type' : d_dependency_types}

d4 = {'Action Caption': a_captions,
      'Action Name': a_names,
      'Activation Typen': a_activation_types,
      'Auto Clear': a_activation_auto_clears,
      'Source Dashboard': a_source_dashboards,
      'Source Worksheet': a_source_worksheets,
      'Source Type': a_source_types,
      'Command Name': a_cmd_param_names,
      'Command Value': a_cmd_param_values}
      
d6 = {'One': m_names, 'Two': m_values}

df1 = DataFrame(d1, columns=['Worksheet Name', 'Field Caption', 'Field Name',
'Role', 'Datatype', 'Types', 'Formula', 'Default Format'])
df2 = DataFrame(d2, columns=['Parameter Caption', 'Parameter Name', 'Datatype',
                             'Domain Type', 'Types', 'Format', 'Aggregation', 
                             'Role', 'Semantic Role', 'Granularity', 'Minumum',
                             'Default Value', 'Maximum', 'Formula'])
df3 = DataFrame(d3, columns=['Dashboard Name',
      'Dashboard Min Width',
      'Dashboard Max Width',
      'Dashboard Min Height',
      'Dashboard Max Height',
      'Datasource Caption', 
      'Datasource Name',
      'Datasource Dependency',
      'Dependency Name', 
      'Dependency Datatype', 
      'Dependency Role',
      'Dependency Type'])

df4 = DataFrame(d4, columns=['Action Caption',
      'Action Name',
      'Activation Typen',
      'Auto Clear',
      'Source Dashboard',
      'Source Worksheet',
      'Source Type',
      'Command Name',
      'Command Value'])
      
df6 = DataFrame(d6)

#Workaround until Field loops are fixed
df1 = df1.drop_duplicates()

#Removing old version of file and creating new one
file_xlsx = os.path.dirname(file_twb) +'\\'+fileName+'.xlsx'

try:
    os.remove(file_xlsx)
except OSError:
    pass

#Saving data to excel file
writer = ExcelWriter(file_xlsx)
df1.to_excel(writer,'Fields', index=False)
df2.to_excel(writer,'Parameters', index=False)
df3.to_excel(writer,'Dashboards', index=False)
df4.to_excel(writer,'Actions', index=False)
df6.to_excel(writer,'Misc', index=False,  header=False)
writer.save()

#Cleaning up
try:
    os.remove(file_xml)
except OSError:
    pass