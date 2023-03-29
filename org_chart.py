import networkx as nx

import matplotlib.pyplot as plt
from matplotlib.offsetbox import AnnotationBbox, OffsetImage

import csv

import PIL
from textwrap3 import wrap

import os

import win32ui

from datetime import datetime

def org_chart():

    config = {
        "node_size" : 1,
        "node_to_edge_distance" : 400,
        "member_name_y_offset" : -.01,
        "member_name_x_alignment" : "center",
        "member_name_y_alignment" : "top",
        "member_name_font_size" : 2,
        "member_name_wrap" : 12,
        "member_name_facecolor":"#6666ff",
        "icon_size" : .05,
        "icon_y_offset" : 0.02,
        "dpi" : 600,
    }
    # csv.DictReader reads from a csv file and creates an object 
    # similar to a list-of-dictionaries.  It uses the first row as keys
    # Example source:
    #   key,reports_to,full_name,icon
    #   bgoldberg,kbinaxas,Benjamin Goldberg,ben.jpg    
    #   kbinaxas,crudy,Kyle Binaxas,kyle.jpg
    reader = csv.DictReader(open('data/data.csv', 'r'))
    # result: an iterable object which contains a list of dictionaries, like this:
    #[
    #  {'key': 'bgoldberg', 'reports_to': 'kbinaxas', 'full_name': 'Benjamin Goldberg', 'icon': 'ben.jpg'}
    #  {'key': 'kbinaxas', 'reports_to': 'crudy', 'full_name': 'Kyle Binaxas', 'icon': 'kyle.jpg'}
    #]
    # 'key' is just the name I gave to the first column. I could have called it 'anything'

    sorted_csv = sorted(reader, key=lambda row: (row['reports_to'], row['key']))
    # sorted sorts and also creates a list.  DictReader, which is a list-like object, wouldn't work with the next line of code

    #Create a dictonary, with each entry being a key and sub-dictionary
    members = { member['key']:member for member in sorted_csv if member['key'] > ""}
    #example:
    #   { 'bgoldberg': {'key':'bgoldberg','friendly_name':'Ben', 'full_name':'Benjamin Goldberg', 'position_name':'Tech Services Coordinator', 'reports_to':'kbinaxas', 'icon':'ben.jpg'}}
    #   { 'kbinaxas':  {'key':'kbinaxas','friendly_name':'Kyle', 'full_name':'Kyle Binaxas', 'position_name':'Tech Services Manager', 'reports_to':'crudy', 'icon':'kyle.jpg'}}
    #the 'key' field in the sub-dictionary is unnecessary but it was easier just to keep it there

    # create the Networkx graph
    G = nx.DiGraph()
    # For all it's crunching, the data in the objects that networkx produces is relatively simple.  Just nodes and edges

    for key,member in members.items():

        G.add_node (
            key,
            full_name = member['full_name'],
            reports_to = member['reports_to'],
            icon = os.path.join('icons', member['icon']),
        )
    # the graph now has a dictionary of nodes similar to the dictionary tha was fed to it

    # In math, an edge is a line that connect nodes.  In networkx, it's just a tuple with two node elements  
    # and an optional dictonary of values. I don't need any dictionaries for my edges
    # but if I wanted to, I could add other fields like I did with add_node    
    for n in G.nodes:
        if G.nodes[n]['reports_to'] is not None and G.nodes[n]['reports_to'] > "":
            G.add_edge(G.nodes[n]['reports_to'], n)

    # topological_generations groups a graph by the amount of ancestors in its connections
    # In my case the boss will be alone in the first group, department heads together in the next group, 
    # peons like me in groups depending on how far we are from the boss

    # I create a new graph and add nodes to the new graph in layer order.  This keeps the edges from 
    # crossing on the final product

    H = nx.DiGraph()
    for i, layer in enumerate(nx.topological_generations(G)):
        for n in layer:
            G.nodes[n]["layer"] = i
            H.add_node(n,**G.nodes[n])

    H.add_edges_from(G.edges(data=True))

    G=H

    #create the layout, which will be a dict keyed to the graph's nodes, and include xy coordinates of each node like:
    #   bgoldberg :  [ 1.         -0.01960784] 
    #   kbinaxas :   [0.33333333   0.31372549]
    pos = nx.multipartite_layout(G, subset_key="layer", align="horizontal")
    # the numbers in each tuple represent x and y and range from about -1 to 1, 


    # reverse the y coordinates because nx.multipartite_layout assigned them opposite of what I want
    for k in pos:
        pos[k][1] *= -1

    # create the figure and subplot
    # this is a matplotlib construct.  A figure is the overall product, and a figure can have multiple subplots
    # I only need one subplot. By convention, I'll refer to the sublot as ax
    fig, ax = plt.subplots()

    #draw the nodes and and edges 
    #these are networkx functions which call on matplotlib.  These place the nodes and edges on the subplot
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=config["node_size"])
    nx.draw_networkx_edges(G, pos, ax=ax, node_size=config["node_to_edge_distance"])
    # for draw_netwokx_nodes, node_size is (as expected) how big to draw the nodes 
    # for draw_netwokx_edges, node_size determines where to place the ends of the edges

    #put the members's names on the plot
    for n in G.nodes:
        # use textwrap3 to create wrapped text.  wrap breaks the text into a list, which I rejoin '\n'
        display_name="\n".join(wrap(G.nodes[n]["full_name"],config["member_name_wrap"]))
        ax.text(
            pos[n][0], 
            pos[n][1]+config["member_name_y_offset"],
            display_name,
            ha=config["member_name_x_alignment"], 
            va=config["member_name_y_alignment"],
            fontsize=config["member_name_font_size"],
            bbox={"facecolor":config["member_name_facecolor"]}
        )

    for n in pos:
        if G.nodes[n]['icon']  > "":
            #I have to handle errors encountered when the image loads or the script will crash
            try:
                icon_image = PIL.Image.open(G.nodes[n]['icon'])
                # AnnotationBbox isn't specific to images.  I could have used it for the text
                ab = AnnotationBbox(
                    OffsetImage(
                        icon_image,
                        zoom=config["icon_size"],
                    ),
                    (pos[n][0], pos[n][1]+config["icon_y_offset"]),
                    frameon=False
                )  
                ax.add_artist(ab)

            except:
                icon_image = None
                print('Failure to load image from node[{}], {}'.format(G.nodes[n],G.nodes[n]['icon'] ))                    
                        
    #this gets rid of the border which is necessary because some of the text goes outside the border     
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)

    dlg = win32ui.CreateFileDialog(0)
    dlg.SetOFNInitialDir('~')
    dlg.DoModal()

    plt.savefig(
        dlg.GetPathName(),
        format='png',
        dpi=config["dpi"]
    )
    # plt.savefig(
    #     # this long line is just to generate a unique name for each output
    #     os.path.join('output','org_chart_{}.png'.format(datetime.strftime(datetime.today(),'%Y%m%d%H%M%S%f'))), 
    #     format='png', 
    #     dpi=config["dpi"]
    # )

org_chart()
