#**Markus**

#####Tool for generating Word documents based on Excel file content. <sup>with gui</sup>

####Prerequisites
+ MS Windows
+ MS Office

###How to use
+ Create docx template with jija2 style variables(`{{ variable name }}`) 
+ Create Excel file where one sheet will contain all necesary information.
 + First line should contain variables names
 + Second line can be description
 + Columns where no variable is set are ignored by Markus
+ Configure markus.conf in the same directory as Markus executable
 + docx_templates_count        - number of templates to process
 + docx_template_name_1        - name of template 1 
 + docx_template_naming_1      - format for resulting document name for template 1 (see usage of variables in examples)
 + docx_template_dir_1         - initial directory for open template dialog for template 1
 + source_file_dir             - initial directory for open Excel file dialog
 + complete_data_sheet_name    - name of sheet in Excel file which should be processed
 + lines_to_skip               - number of lines used for description(yellow color in examples)
 + var_to_indicate_row         - variable to indicate row as filled 