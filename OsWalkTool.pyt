# -*- coding: utf-8 -*-

import arcpy
import os #Used to find the templates used in combination with this tool
import re #Used to manipulate filenames with regex
import pandas as pd #Used to write to excel
import datetime #Used to get the last modified time of a file

#6,159 total .mxds -- need to stop using Excel as a place to hold data before data summary is generated.


#aprx = arcpy.mp.ArcGISProject(template_path)
#dummy = aprx.saveACopy(r"C:\bug_zone\mxd_processing\mxd_processing_template\mxd_processing_in_progress.aprx")  #Paramaterize this later
#mxd_path = (r"C:\bug_zone\Florida\SW_Peninsular_Florida_ESI_2016_10-3.mxd") #Paramaterize this later
#outfile = r"\bug_zone\mxd_processing\mxd_processing_template\test.xlsx"
#dir_to_walk = r"C:\Users\meyerr4224\OneDrive - ARCADIS\Desktop\advanced_testing\_ArcGISPro_Templates\Layers"



class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Data source finder"
        self.alias = "Data source finder"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Data source finder"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):

        """Define parameter definitions"""
        folder_to_walk = arcpy.Parameter(
            displayName = "Select a folder to crawl through to find files", 
            name = "infolder", 
            datatype = 'DEFolder', 
            direction = 'Input', 
            parameterType = 'Required', 
        )

        output_folder = arcpy.Parameter(
            displayName = "Select the folder to output the excel file to", 
            name = "outfolder", 
            datatype = 'DEFolder', 
            direction = 'Input', 
            parameterType = 'Required', 
        )

        output_file_name = arcpy.Parameter(
            displayName = "Select the output file name (include .xlsx)", 
            name = "outfile_name", 
            datatype = 'String', 
            direction = 'Input', 
            parameterType = 'Required', 
        )

        params = [folder_to_walk, output_folder, output_file_name]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This hmethod is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """Executes when tool is run"""
        def p_addmessage(msg):
            print(msg)
            arcpy.AddMessage(msg)

        """CONSTANTS"""
        outfile = os.path.join(parameters[1].valueAsText, parameters[2].valueAsText) #Output folder + output filename

        #This script relies on the existence of a template aprx file that is used to process .mxds, .lyrx, and .lyr files.
        #Therefore, an internal variable is defined to point to the location of the template aprx file. 
        #By default this is in the same folder as the script, hence the below.
        local_path = os.path.dirname(os.path.realpath(__file__))
        template_path = os.path.join(local_path, "mxd_processing_template.aprx")
        dummy_path = os.path.join(local_path, "mxd_processing_in_progress.aprx")

  

        """FUNCTION DEFINITIONS"""
        def get_parent_folder_mtime(file):
            parent_folder = os.path.dirname(file)
            mtime = os.path.getmtime(parent_folder)
            #turn mtime into datetime object
            mtime = datetime.datetime.fromtimestamp(mtime)
            #turn mtime into string
            mtime = mtime.strftime("%m/%d/%Y")
            return mtime
        

        def parse_mxd(in_mxd, template_aprx):
            layer_objs = [] #List of layer objects used to populate excel sheet
             #To handle mxds, we import them into a dummy project that lets us inspect their metadata.
            template = arcpy.mp.ArcGISProject(template_aprx)
            file_name = in_mxd.split("\\")[-1]

            dummy = template.saveACopy(dummy_path)
            d_aprx = arcpy.mp.ArcGISProject(dummy_path)
            d_aprx.importDocument(in_mxd)
            all_maps = d_aprx.listMaps()
            mtime = get_parent_folder_mtime(in_mxd)
            for map in all_maps:
                for layer in map.listLayers():
                    
                    layer_obj = {
                    "file_name": file_name,
                    "map": map.metadata.title,
                    "name": None,
                    "data_source": None,
                    "title": None,
                    "summary": None,
                    "description": None,
                    "tags": None,
                    "folder_modified": mtime,
                    "is_broken": layer.isBroken
                    }

                    if layer.supports("dataSource"):
                        pattern = re.compile(r'<[^>]+>') #Remove HTML from descriptions.
                        layer_obj['name'] = layer.name
                        layer_obj['data_source'] = layer.dataSource
                        layer_obj['title'] = layer.metadata.title
                        layer_obj['summary'] = layer.metadata.summary
                        layer_obj['tags'] = layer.metadata.tags
                        if(layer.metadata.description):
                            desc_with_html = layer.metadata.description
                            clean_desc = re.sub(pattern, '', desc_with_html)
                            layer_obj['description'] = clean_desc
                        layer_objs.append(layer_obj) 

                #TODO: Delete the dummy APRX when done. Os delete at dummy_path
                
                return layer_objs
       
        def parse_aprx(in_aprx: str):
            layer_objs = [] #List of layer objects used to populate excel sheet
            aprx = arcpy.mp.ArcGISProject(in_aprx) #Expects path to an aprx
            file_name = in_aprx.split("\\")[-1]
            if file_name=="mxd_processing_template.aprx" or file_name=="mxd_processing_in_progress.aprx":
                print("Skipping template APRX")
                return
            all_maps = aprx.listMaps()
            mtime=get_parent_folder_mtime(in_aprx)
            for map in all_maps:
                for layer in map.listLayers():

                    layer_obj = {
                    "file_name": file_name,
                    "map": map.metadata.title,
                    "name": None,
                    "data_source": None,
                    "title": None,
                    "summary": None,
                    "description": None,
                    "tags": None,
                    "folder_modified": mtime,
                    "is_broken": layer.isBroken
                    }

                    if layer.supports("dataSource"):
                        pattern = re.compile(r'<[^>]+>') #Remove HTML from descriptions.
                        #print(inspect.getmembers(layer.metadata)) #May need to remove the final part of the path...Can these nest?
                        layer_obj['name'] = layer.name
                        layer_obj['data_source'] = layer.dataSource
                        layer_obj['title'] = layer.metadata.title
                        layer_obj['summary'] = layer.metadata.summary
                        layer_obj['tags'] = layer.metadata.tags
                        if(layer.metadata.description):
                            desc_with_html = layer.metadata.description
                            clean_desc = re.sub(pattern, '', desc_with_html)
                            layer_obj['description'] = clean_desc
                        layer_objs.append(layer_obj) 

                #TODO: Delete the dummy APRX when done. Os delete at dummy_path
                
                return layer_objs

        def parse_lyrx(in_lyrx: str):
            lyr_file = arcpy.mp.LayerFile(in_lyrx)
            lyr_file_name = in_lyrx.split("\\")[-1]
            ##Shorten to 31 characters if longer. Excel sheet name limitation.
            lyr_file_name = lyr_file_name[:31]
            layer_objs = [] #List of layer objects used to populate excel sheet
            mtime = get_parent_folder_mtime(in_lyrx)
            for lyr in lyr_file.listLayers():
                    
                    layer_obj = {
                    "file_name": lyr_file_name,
                    #"map": map.metadata.title,
                    "name": None,
                    "data_source": None,
                    "title": None,
                    "summary": None,
                    "description": None,
                    "tags": None,
                    "folder_modified": mtime,
                    "is_broken": lyr.isBroken
                    }

                    if lyr.supports("dataSource"):
                        pattern = re.compile(r'<[^>]+>') #Remove HTML from descriptions.
                        layer_obj['name'] = lyr.name
                        layer_obj['data_source'] = lyr.dataSource
                        layer_obj['title'] = lyr.metadata.title
                        layer_obj['summary'] = lyr.metadata.summary
                        layer_obj['tags'] = lyr.metadata.tags
                        if(lyr.metadata.description):
                            desc_with_html = lyr.metadata.description
                            clean_desc = re.sub(pattern, '', desc_with_html)
                            layer_obj['description'] = clean_desc
                        layer_objs.append(layer_obj) 

                    #TODO: Delete the dummy APRX when done. Os delete at dummy_path
                
            return layer_objs

        def parse_lyr(in_lyr: str):
            lyr_file = arcpy.mp.LayerFile(in_lyr)
            lyr_file_name = in_lyr.split("\\")[-1]
            ##Shorten to 31 characters if longer. Excel sheet name limitation.
            lyr_file_name = lyr_file_name[:31]
            layer_objs = []
            mtime = get_parent_folder_mtime(in_lyr)

            for lyr in lyr_file.listLayers():
                    
                    layer_obj = {
                    "file_name": lyr_file_name,
                    #"map": map.metadata.title,
                    "name": None,
                    "data_source": None,
                    "title": None,
                    "summary": None,
                    "description": None,
                    "tags": None,
                    "folder_modified": mtime,
                    "is_broken": lyr.isBroken
                    }

                    if lyr.supports("dataSource"):
                        pattern = re.compile(r'<[^>]+>') #Remove HTML from descriptions.
                        layer_obj['name'] = lyr.name
                        layer_obj['data_source'] = lyr.dataSource
                        layer_obj['title'] = lyr.metadata.title
                        layer_obj['summary'] = lyr.metadata.summary
                        layer_obj['tags'] = lyr.metadata.tags
                        if(lyr.metadata.description):
                            desc_with_html = lyr.metadata.description
                            clean_desc = re.sub(pattern, '', desc_with_html)
                            layer_obj['description'] = clean_desc
                        layer_objs.append(layer_obj) 

                    #TODO: Delete the dummy APRX when done. Os delete at dummy_path
                
            return layer_objs


        def write_to_excel(in_dict, outfile, sheetname):
            #Check if an outfile exists at the path. If it does not, create one + sheet with headers.
            layer_df = pd.DataFrame(in_dict)
            if not os.path.isfile(outfile):
                writer = pd.ExcelWriter(outfile, engine='openpyxl')
                layer_df.to_excel(writer, sheet_name=sheetname, index=False)
                writer.save()

            #Check if the sheet exists. If not, create it with headers.
            if not sheetname in pd.ExcelFile(outfile).sheet_names:
                writer = pd.ExcelWriter(outfile, engine='openpyxl', mode='a')
                layer_df.to_excel(writer, sheet_name=sheetname, index=False)
                writer.save()
                return
            
            #Otherwise, append to extant sheet with no headers
            #TODO: NOTE: The way this is set up is that if the file encounters multiple
            existing_data = pd.read_excel(outfile, sheet_name=sheetname)
            writer = pd.ExcelWriter(outfile, engine="openpyxl", mode='a', if_sheet_exists="overlay")
            layer_df.to_excel(writer, sheet_name=sheetname, startrow=writer.sheets[sheetname].max_row, header=False, index=False)
            writer.save()
            return


        def summary_from_data_frame(_all_layers, _outfile):
            output_df = pd.DataFrame(columns=['data_source', 'requested_by']) #May not even need 'data_source' here if you're being like this.
            unique_sources = _all_layers['data_source'].unique()
            for source in unique_sources:
                #Maybe this needs to exclude the .aprx and .mxds? Currently their inclusion makes this off by one per each.
                count = _all_layers['data_source'].value_counts()[source]
                #Write the source and count to the summary_by_source sheet
                df = pd.DataFrame(columns=['data_source', 'num_requests', 'requested_by', 'num_unique_requesters', 'broken_over_total'])
                df['data_source'] = [source]
                df['num_requests'] = [count] #Num_requests is number of LAYERS within the projects/maps/etc. that request a source.
                #Every unique source gets a list of all the projects that use it. Projects are defined by being .aprx or .mxd.
                _requested_by = [_all_layers.loc[_all_layers['data_source'] == source]['file_name'].unique()] #Unique projects/maps/etc. that do the requesting
                _requested_series = pd.Series(_requested_by)
               
                #Create a new series called _filtered_series based on _requested series by filtering only items that end with .aprx or .mxd
                _filtered_series = _requested_series.apply(lambda x: [i for i in x if i.endswith('.aprx') or i.endswith('.mxd')]) #Please explain this to me.   
                df['requested_by'] = _filtered_series
                df['num_unique_requesters'] = [len(_filtered_series[0])]

                #Broken works
                _num_broken = len(_all_layers.loc[(_all_layers['data_source'] == source) & (_all_layers['is_broken'] == True)])
                _broken_over_total = str(_num_broken) + r'/' + str(count)
                df['broken_over_total'] = _broken_over_total
                output_df = pd.concat([output_df, df], join='outer', ignore_index=True)

            writer = pd.ExcelWriter(_outfile, engine='openpyxl', mode='w')    
            output_df.to_excel(writer, sheet_name='summary_by_source', index=False)
            writer.save()
            return



        """END FUNCTION DEFINITIONS"""

        """MAIN EXECUTION"""

        
        #Large dataframe that holds all layers of all items found in the target path.
        all_layers = pd.DataFrame(columns=['file_path', 'file_name', 'map', 'name', 'data_source', 'title', 'summary', 'description', 'tags', 'folder_modified'])

        
        def walk_and_exec_df(directory, _all_layers):
            layers_to_return = _all_layers.copy(deep=True) #just making sure we don't do any modifying in place
            for root, dirs, files in os.walk(directory):
                for file in files:
                    if file.endswith(".mxd"):
                        mxd_path = os.path.join(root, file)
                        mxd_data = parse_mxd(mxd_path, template_path)
                        layers_to_return = pd.concat([layers_to_return, pd.DataFrame(mxd_data)], ignore_index=True)
                        print ("DONE EXECUTING PARSE_MXD")
                    elif file.endswith(".aprx"):
                        aprx_path = os.path.join(root, file)
                        aprx_data = parse_aprx(aprx_path)
                        layers_to_return = pd.concat([layers_to_return, pd.DataFrame(aprx_data)], ignore_index=True)
                    elif file.endswith(".lyrx"):
                        lyrx_path = os.path.join(root, file)
                        lyrx_data = parse_lyrx(lyrx_path)
                        layers_to_return = pd.concat([layers_to_return, pd.DataFrame(lyrx_data)], ignore_index=True)
                    elif file.endswith(".lyr"):
                        lyr_path = os.path.join(root, file)
                        lyr_data = parse_lyr(lyr_path)
                        layers_to_return = pd.concat([layers_to_return, pd.DataFrame(lyr_data)], ignore_index=True)
                    else:
                        print("Skipping file: " + file)

        
            return layers_to_return
        
        output = walk_and_exec_df(parameters[0].valueAsText, all_layers) #Directory to walk
        summary_from_data_frame(output, outfile)


        #TODO: What on earth do we do with access limitations? Add permissions check to walk_and_exec_excel? Or just give to a super user?


        """END MAIN EXECUTION"""

        #ArcGIS Pro return below#
        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
