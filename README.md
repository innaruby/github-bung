{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 1,
   "id": "1c37c3d9-0eed-4203-9310-5d8d63a76d19",
   "metadata": {
    "hide_input": false,
    "jupyter": {
     "source_hidden": true
    }
   },
   "outputs": [],
   "source": [
    "#importing required moduls\n",
    "import pandas as pd\n",
    "import ipywidgets as widgets\n",
    "from IPython.display import display\n",
    "import traceback\n",
    "from processing import open_an_excel, get_msr_output_format, get_product_MSR_pos_alignment, create_report, update_target_excel_xlwings, custom_formatter"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "602919d4-1dc3-4dcb-8d42-2b33b8fdc494",
   "metadata": {
    "hide_input": true
   },
   "source": [
    "## User Instructions\n",
    "\n",
    "#### 1. Click on \"Cell\" -> \"Run all\" in the above menu to show User Interface\n",
    "#### 2. Click on the \"Get support file\"  button below -> select the excel file in the dialog box with the product tree & the MSR - cost center assignment sheets\n",
    "#### 3. Select MSR report output format in the dropdown\n",
    "#### 4 Check/uncheck applicable cost centers via the checkboxes\n",
    "#### 5. Click on \"Get Data File\" button below -> select the excel file in the dialog box with the KUKA report\n",
    "#### 6. Optionally change if you want to see a summary or broken down to cost centers"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 2,
   "id": "ccd66662-40f0-45f3-9dca-10fafafc9af3",
   "metadata": {
    "hide_input": false,
    "jupyter": {
     "source_hidden": true
    },
    "scrolled": true
   },
   "outputs": [
    {
     "data": {
      "application/vnd.jupyter.widget-view+json": {
       "model_id": "e2b518afa53f46efbeefffa37af5b8dc",
       "version_major": 2,
       "version_minor": 0
      },
      "text/plain": [
       "VBox(children=(HBox(children=(VBox(children=(Button(description='Get Input File', style=ButtonStyle()),), layo…"
      ]
     },
     "metadata": {},
     "output_type": "display_data"
    }
   ],
   "source": [
    "#----------------------------------------------------------------------\n",
    "# 1) CREATE WIDGETS + OUTPUT\n",
    "# ---------------------------------------------------------------------\n",
    "\n",
    "# 1. the input file buttons\n",
    "#supp_file_button = widgets.Button(description=\"Get Support File\")\n",
    "input_file_button = widgets.Button(description=\"Get Input File\")\n",
    "recalculate_button = widgets.Button(description=\"Recalculate\", layout=widgets.Layout(margin='1'))\n",
    "\n",
    "element_layout = widgets.Layout(\n",
    "    width='100%',\n",
    "    min_width='90px',\n",
    "    margin ='0px 0px 0px 0px'\n",
    ")\n",
    "\n",
    "# 2. the selectors\n",
    "gesamt_sheet_selector = widgets.Dropdown(options=[],layout=element_layout)\n",
    "\n",
    "fg_sheet_selector = widgets.Dropdown(options=[],layout=element_layout)\n",
    "\n",
    "format_selector = widgets.Dropdown(options=[], layout=element_layout)\n",
    "\n",
    "report_type_selector = widgets.ToggleButtons(\n",
    "    options=['Summary', 'Detailed'],\n",
    "    button_style='',\n",
    "    tooltips=['All the cost centers will be summed together to one column', 'The Positions will be presented parallel for each cost center'],\n",
    ")\n",
    "\n",
    "# 3. the incuded cost centers\n",
    "checkboxes_container = widgets.HBox(layout=widgets.Layout(display='flex', justify_content='flex-start',align_content='flex-start'))\n",
    "\n",
    "# 4. the section for the target MSR file\n",
    "target_column_input = widgets.Text(\n",
    "    value= pd.Timestamp.now().date().strftime('%m/%Y'),\n",
    "    disabled=False,\n",
    "    layout=element_layout,\n",
    "    )\n",
    "\n",
    "date_row_input = widgets.IntText(value= 8,disabled=False,layout=element_layout)\n",
    "\n",
    "data_row_input = widgets.IntText(value= 115,disabled=False,layout=element_layout)\n",
    "\n",
    "target_file_button = widgets.Button(description=\"Update MSR File\")\n",
    "\n",
    "data_output_area = widgets.Output()\n",
    "export_output_area = widgets.Output()\n",
    "\n",
    "# ---------------------------------------------------------------------\n",
    "# 2) FUNCTIONS TO CALL ON WIDGET CHANGES\n",
    "# ---------------------------------------------------------------------\n",
    "# 1. Function that reads in the files and updates the format selector widget\n",
    "def input_file_button_clicked(b):\n",
    "    try: \n",
    "        with data_output_area:\n",
    "            data_output_area.clear_output()\n",
    "            \n",
    "            input_xl = open_an_excel(titlestr='Browse for the ProduktBaum/Kostenstelle file')\n",
    "            if not input_xl:\n",
    "                print('No input file selected')\n",
    "                gesamt_sheet_selector.options = []\n",
    "                fg_sheet_selector.options = []\n",
    "                format_selector.options = []\n",
    "                input_file_button.input_xl = None\n",
    "            else:\n",
    "                input_file_button.input_xl = input_xl\n",
    "                print('Report sheet read in...')\n",
    "            \n",
    "                current_pt = get_product_MSR_pos_alignment(input_xl, 'Produktbaum')\n",
    "                if current_pt:\n",
    "                    input_file_button.producttree = current_pt\n",
    "                    print('Product tree read in...')\n",
    "                else:\n",
    "                    input_file_button.producttree = None\n",
    "            \n",
    "                local_ccs = get_msr_output_format(input_xl,\"Kostenstellen\")\n",
    "                if local_ccs: \n",
    "                    format_selector.ccs = local_ccs\n",
    "                    format_selector.options = list(local_ccs.keys())\n",
    "                    format_selector.value = format_selector.options[0]\n",
    "                    print('MSR-Cost centers read in...')\n",
    "                else:\n",
    "                    format_selector.ccs = local_ccs = None\n",
    "                    format_selector.options = []\n",
    "                    format_selector.value = None\n",
    "            \n",
    "                gesamt_sheet_selector.options = [sheet for sheet in input_xl.sheet_names if \"KUKA-Gesamt\" in sheet]\n",
    "                if len(gesamt_sheet_selector.options)==0:\n",
    "                    print('No sheet with the expression \"KUKA-Gesamt\" in the sheet name in file')\n",
    "                    gesamt_sheet_selector.value = None\n",
    "                else:\n",
    "                    gesamt_sheet_selector.value = gesamt_sheet_selector.options[-1]\n",
    "            \n",
    "                fg_sheet_selector.options = [sheet for sheet in input_xl.sheet_names if \"KUKA-Finanz\" in sheet]\n",
    "                if len(fg_sheet_selector.options)==0:\n",
    "                    print('No sheet with the expression \"KUKA-Finanz\" in the sheet name in file')\n",
    "                    fg_sheet_selector.value = None\n",
    "                else:\n",
    "                    fg_sheet_selector.value = fg_sheet_selector.options[-1]\n",
    "\n",
    "            generate_report()\n",
    "    except Exception as e:\n",
    "        with data_output_area:\n",
    "            data_output_area.clear_output()\n",
    "            print(\"Error reading the input file:\", e, traceback.format_exc())\n",
    "\n",
    "# 2. Function that regenerates checkboxes on format selector dropdown change\n",
    "def regenerate_checkboxes(change):\n",
    "    \"\"\"Rebuild checkboxes when the selected dropdown value changes.\"\"\"\n",
    "    # Only act if the 'value' changed\n",
    "    if change['type'] == 'change' and change['name'] == 'value':\n",
    "        if change['new'] is None:\n",
    "            checkboxes_container.children = []\n",
    "            return\n",
    "        selected_key = change['new']\n",
    "        \n",
    "        # Retrieve ccs from the widget attribute\n",
    "        local_ccs = format_selector.ccs\n",
    "        \n",
    "        # Build a new list of checkboxes for the selected key\n",
    "        new_checkboxes = []\n",
    "        for val in local_ccs[selected_key]:\n",
    "            cb = widgets.Checkbox(\n",
    "                value=True,\n",
    "                description=str(val),\n",
    "                indent=False\n",
    "            )\n",
    "            new_checkboxes.append(cb)\n",
    "        \n",
    "        # Replace old checkboxes in the container with the new ones\n",
    "        checkboxes_container.children = new_checkboxes\n",
    "\n",
    "def get_checked_checkbox_list():\n",
    "    checked_values = []\n",
    "    for cb in checkboxes_container.children:\n",
    "        if cb.value:  # if checkbox is checked\n",
    "            checked_values.append(int(cb.description))\n",
    "    return checked_values\n",
    "\n",
    "def on_target_file_button_clicked(b):\n",
    "    try:\n",
    "        if not hasattr(data_output_area, 'result'):\n",
    "            print(\"No data has been loaded yet.\")\n",
    "            return\n",
    "        data_input = data_output_area.result\n",
    "        if isinstance(data_input,pd.Series):\n",
    "            used_cost_centers = [data_input.name]\n",
    "        elif isinstance(data_input,pd.DataFrame):\n",
    "            used_cost_centers = data_input.columns.tolist()\n",
    "        else:\n",
    "            print(\"No proper data available.\")\n",
    "            return\n",
    "        with export_output_area:\n",
    "            export_output_area.clear_output()\n",
    "            update_target_excel_xlwings(\n",
    "                data_input=data_input,\n",
    "                cost_centers_used=used_cost_centers,\n",
    "                target_column=target_column_input.value,\n",
    "                date_row=date_row_input.value,\n",
    "                data_row=data_row_input.value,\n",
    "                titlestr='Please select the MSR file to insert the data into'\n",
    "            )\n",
    "        \n",
    "    except Exception as e:\n",
    "        with export_output_area:\n",
    "            export_output_area.clear_output()\n",
    "            print(\"Error exporting the data into the MSR file:\", e, traceback.format_exc())\n",
    "            \n",
    "def on_report_type_change(change):\n",
    "    if change[\"name\"] == \"value\":\n",
    "        generate_report()\n",
    "\n",
    "def on_recalc_button_clicked(b):\n",
    "    generate_report()\n",
    "    \n",
    "def generate_report():\n",
    "    with data_output_area:\n",
    "        data_output_area.clear_output()\n",
    "        try:\n",
    "            # If user hasn't loaded a file yet, there's nothing to process\n",
    "            if not hasattr(input_file_button, 'input_xl'):\n",
    "                print(\"No data file loaded yet.\")\n",
    "                return\n",
    "            \n",
    "            # \"report_xl\" is the Excel file stored\n",
    "            report_xl = input_file_button.input_xl\n",
    "            \n",
    "            product_tree = getattr(input_file_button, 'producttree', None)\n",
    "            if product_tree is None:\n",
    "                print(\"No product tree loaded yet.\")\n",
    "                return\n",
    "            \n",
    "            used_cost_centers = get_checked_checkbox_list()\n",
    "            \n",
    "            report_type = report_type_selector.value\n",
    "            \n",
    "            result = create_report(\n",
    "                report_xl=report_xl,\n",
    "                gesamt_sheet_name=gesamt_sheet_selector.value,\n",
    "                fg_sheet_name=fg_sheet_selector.value,\n",
    "                current_pt=product_tree,\n",
    "                used_cost_centers=used_cost_centers,\n",
    "                #header_row=9,\n",
    "                #cost_center_colname='Kostenstelle des Geschäfts',\n",
    "                #sachkonto_colname='Sachkonto-Nr.',\n",
    "                report_type=report_type,\n",
    "                #main_groups_list:list[str]=['Aktiv','Leasing','Factoring','Bankverbindl.','Giro','Spar'],\n",
    "                #sachkonto_details_for:list[str]=['Giro']\n",
    "            )\n",
    "            data_output_area.result = result\n",
    "            if isinstance(result,pd.Series):\n",
    "                result_to_show = result.to_frame()\n",
    "            else:\n",
    "                result_to_show = result\n",
    "            \n",
    "            if result_to_show is None:\n",
    "                print('No data was returned...')\n",
    "                return\n",
    "            # Show the styled DataFrame (HTML) in the output widget\n",
    "            display(result_to_show.style.format(custom_formatter))\n",
    "            \n",
    "            \n",
    "        except Exception as e:\n",
    "            print(\"ERROR processing data:\", e, traceback.format_exc())\n",
    "            \n",
    "# ---------------------------------------------------------------------\n",
    "# 3) OBSERVER FUNCTIONS\n",
    "# ---------------------------------------------------------------------\n",
    "format_selector.observe(regenerate_checkboxes)\n",
    "report_type_selector.observe(on_report_type_change, names='value')\n",
    "input_file_button.on_click(input_file_button_clicked)\n",
    "#supp_file_button.on_click(on_supp_file_button_clicked)\n",
    "#report_file_button.on_click(on_data_file_button_clicked)\n",
    "target_file_button.on_click(on_target_file_button_clicked)\n",
    "recalculate_button.on_click(on_recalc_button_clicked)\n",
    "# ---------------------------------------------------------------------\n",
    "# 4) DISPLAYING THE WIDGETS\n",
    "# ---------------------------------------------------------------------\n",
    "\n",
    "vbox_layout = widgets.Layout(\n",
    "    width='25%',\n",
    "    min_width='100px',\n",
    "    margin ='0px 10px 2px 10px',\n",
    ")\n",
    "\n",
    "ui = widgets.VBox([\n",
    "    widgets.HBox([\n",
    "        widgets.VBox([input_file_button],layout=widgets.Layout(justify_content='flex-end',margin='0px')),\n",
    "        widgets.VBox([widgets.Label('KUKA-Gesamtreport sheet:'),gesamt_sheet_selector],layout=vbox_layout),\n",
    "        widgets.VBox([widgets.Label('KUKA-Finanzgeschäftreport sheet:'),fg_sheet_selector],layout=vbox_layout),\n",
    "        widgets.VBox([widgets.Label('MSR report format:'),format_selector],layout=vbox_layout),\n",
    "    ]),    \n",
    "    checkboxes_container,\n",
    "    widgets.HBox([\n",
    "        widgets.VBox([recalculate_button],layout=widgets.Layout(justify_content='flex-end',margin='0px')),\n",
    "        widgets.VBox([\n",
    "             widgets.Label('Report type:'), report_type_selector],\n",
    "                 layout=widgets.Layout(\n",
    "                     width='75%',\n",
    "                     min_width='300px',\n",
    "                     margin ='0px 10px 2px 10px')),\n",
    "    ]),\n",
    "    \n",
    "])\n",
    "display(ui)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "id": "abd82f83",
   "metadata": {
    "scrolled": false
   },
   "outputs": [
    {
     "data": {
      "application/vnd.jupyter.widget-view+json": {
       "model_id": "edfaf1a29c694976a913a3d1218e3976",
       "version_major": 2,
       "version_minor": 0
      },
      "text/plain": [
       "Output()"
      ]
     },
     "metadata": {},
     "output_type": "display_data"
    }
   ],
   "source": [
    "display(data_output_area)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 4,
   "id": "e9275676",
   "metadata": {},
   "outputs": [
    {
     "data": {
      "application/vnd.jupyter.widget-view+json": {
       "model_id": "d35447ac65ce4222bf157a3bddd371b0",
       "version_major": 2,
       "version_minor": 0
      },
      "text/plain": [
       "HBox(children=(VBox(children=(Button(description='Update MSR File', style=ButtonStyle()),), layout=Layout(just…"
      ]
     },
     "metadata": {},
     "output_type": "display_data"
    }
   ],
   "source": [
    "update_ui = widgets.HBox([\n",
    "    widgets.VBox([target_file_button],layout=widgets.Layout(justify_content='flex-end',margin='0px')),\n",
    "    widgets.VBox([widgets.Label('Column in target file:'),target_column_input],layout=vbox_layout),\n",
    "    widgets.VBox([widgets.Label('Date row in sheets of target:'),date_row_input],layout=vbox_layout),\n",
    "    widgets.VBox([widgets.Label('Row where the pasting should start:'),data_row_input],layout=vbox_layout),\n",
    "    ])\n",
    "\n",
    "\n",
    "display(update_ui)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 5,
   "id": "fde4fa31",
   "metadata": {},
   "outputs": [
    {
     "data": {
      "application/vnd.jupyter.widget-view+json": {
       "model_id": "ca46a75eb9f7496ba8b8e7b1cd0bf694",
       "version_major": 2,
       "version_minor": 0
      },
      "text/plain": [
       "Output()"
      ]
     },
     "metadata": {},
     "output_type": "display_data"
    }
   ],
   "source": [
    "display(export_output_area)"
   ]
  }
 ],
 "metadata": {
  "hide_input": true,
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.10.16"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
