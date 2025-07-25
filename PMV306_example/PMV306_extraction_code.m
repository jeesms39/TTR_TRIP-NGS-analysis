clear

% Define input and output file names
PMV_sheet = "PMV_extraction_Seq.xlsx"; % modify the input excel file if requied
filename_fastq = "S11_L001_R1_001.fastq"; % modify the fastq file name here
Output_filename = 'PMV306_count.xlsx'; % the results will be saved in PMV306_count.xlsx 

% Set up options for importing data from the PMV_sheet and read the data into the table
opts = spreadsheetImportOptions("NumVariables", 2);
opts.Sheet = "PMV306"; % Specify the sheet name here
opts.DataRange = "A2:B9"; % Modify the DataRange to match with your gene list
opts.VariableNames = ["GeneID", "Gene"];
opts.VariableTypes = ["string", "string"];
opts = setvaropts(opts, ["GeneID", "Gene"], "WhitespaceRule", "preserve");
opts = setvaropts(opts, ["GeneID", "Gene"], "EmptyFieldRule", "auto");
PMV_extraction_Seq = readtable(PMV_sheet, opts, "UseExcel", false);
clear opts

% Prepare the output table with appropriate variable names
PMV306_gene_name = string(PMV_extraction_Seq.GeneID(1:end-2));
PMV306_gene = string(PMV_extraction_Seq.Gene(1:end-2));
PMV306_V1 = string(PMV_extraction_Seq.Gene(end-1)); % V1 is saved as second last entry
PMV306_V2 = string(PMV_extraction_Seq.Gene(end)); % V2 is saved as last entry

% Initialize the output table with the gene names
PMV306_Gene_table = table([PMV306_gene_name,PMV306_gene]);
PMV306_Gene_table = splitvars(PMV306_Gene_table);
PMV306_Gene_table.Properties.VariableNames = {'Gene ID','Gene'};

% Initialize the count results for each gene
Count_PMV306 = zeros(length(PMV306_gene), 1);

% Read the FASTQ file and extract sequences
read_sequ= fastqread(filename_fastq);
temp = struct2cell(read_sequ.').'; 
[m1,~] = size(temp);
Sequ = temp(1:m1,2);

% Count occurrences of each gene in the sequences
for n1= 1:length(PMV306_gene)
    clear TF
    for k=1:m1 % Loop through each sequence
        clear Sequ1
        Sequ1= string(Sequ(k));
        % Check if the sequence contains the gene and the two variants
        TF(k) = (contains(Sequ1,PMV306_gene(n1))&contains(Sequ1,PMV306_V1)&contains(Sequ1,PMV306_V2));
        if n1==length(PMV306_gene)
            TF(k) = contains(Sequ1,PMV306_gene(n1)); % Special case for the last gene
        end
    end
    TF = double(TF); % Convert logical array to double
    Count_PMV306(n1) = sum(TF); % Count occurrences
end

% Convert count results to a table and update the gene table
Count_PMV306 = table(Count_PMV306);
Count_PMV306.Properties.VariableNames = {'Count'};
PMV306_Gene_table = [PMV306_Gene_table, Count_PMV306];

% Write the final gene table with counts to the output file
writetable(PMV306_Gene_table,Output_filename,'Sheet','PMV306')