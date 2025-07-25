clear
clc

% Define the input Excel file for data extraction, input FASTQ file and output filename 
PLJR_sheet = "PLJR_extraction_Seq.xlsx"; % modify the input excel file if requied
filename_fastq = "PLJR_Sample_file.fastq"; % modify the fastq file name here
Output_filename = 'PLJR_count.xlsx'; % the results will be saved in PLJR_count.xlsx 

% Set up options for importing data from a spreadsheet
opts = spreadsheetImportOptions("NumVariables", 2);
opts.Sheet = "PLJR"; % Specify the sheet name to import data from
opts.DataRange = "A2:B18"; % Modify the DataRange to match with your gene list
opts.VariableNames = ["GeneID", "Gene"]; 
opts.VariableTypes = ["string", "string"]; 
opts = setvaropts(opts, ["GeneID", "Gene"], "WhitespaceRule", "preserve"); 
opts = setvaropts(opts, ["GeneID", "Gene"], "EmptyFieldRule", "auto"); 
% Read the data from the specified Excel file into a table
PLJR_extraction_Seq = readtable("PLJR_extraction_Seq.xlsx", opts, "UseExcel", false);
clear opts % Clear the options variable to free up memory

% Extract gene names V1 nad V2 from the PLJR_extraction_Seq structure
PLJR_gene_name = string(PLJR_extraction_Seq.GeneID(1:end-2));
PLJR_gene = string(PLJR_extraction_Seq.Gene(1:end-2));
PLJR_V1 = string(PLJR_extraction_Seq.Gene(end-1)); % in the extraction_Seq file V1 is saved as second last entry
PLJR_V2 = string(PLJR_extraction_Seq.Gene(end)); % in the extraction_Seq file V2 is saved as last entry

% Create a table from gene names and gene data
PLJR_Gene_table = table([PLJR_gene_name,PLJR_gene]);
PLJR_Gene_table = splitvars(PLJR_Gene_table);
PLJR_Gene_table.Properties.VariableNames = {'Gene ID','Gene'};

% Initialize a count array for the number of genes
Count_PLJR = zeros(length(PLJR_gene),1);

% Read sequences from a FASTQ file and extract the sequences
read_sequ= fastqread(filename_fastq);
temp = struct2cell(read_sequ.').'; 
[m1,~] = size(temp);
Sequ = temp(1:m1,2);

% Iterate over each gene in the PLJR_gene array and over each sequence
for n1= 1:length(PLJR_gene)
    clear TF 
    for k=1:m1
        clear Sequ1 
        Sequ1= string(Sequ(k)); 
        % Check if the sequence contains the current gene V1 and V2
        TF(k) = (contains(Sequ1,PLJR_gene(n1))&contains(Sequ1,PLJR_V1)&contains(Sequ1,PLJR_V2));
        if n1==length(PLJR_gene)
            TF(k) = contains(Sequ1,PLJR_gene(n1));
        end
    end
    TF = double(TF); % Convert logical array to double
    Count_PLJR(n1) = sum(TF); % Count occurrences of the gene
end

% Convert Count_PLJR to a table format and concatenate to PLJR_Gene_table
Count_PLJR = table(Count_PLJR);
PLJR_Gene_table = [PLJR_Gene_table,Count_PLJR];

% Write the updated table to an Excel file in the specified sheet
writetable(PLJR_Gene_table,Output_filename,'Sheet','PLJR')