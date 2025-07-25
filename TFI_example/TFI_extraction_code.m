clc
clear
% Define input and output file names
TFI_sheet = "TFI_extraction_Seq.xlsx"; % modify the input excel file with  here
Output_filename = 'TFI_count.xlsx'; % the results will be saved in TFI_count.xlsx 
filename_fastq = "TFI_Sample_file.fastq"; % modify the fastq file name here

% Set up options for importing data from the spreadsheet
opts = spreadsheetImportOptions("NumVariables", 4);
opts.Sheet = "TFI";
opts.DataRange = "A2:D209"; % Specify sheet and range
opts.VariableNames = ["pEXCFbackbone_V1", "V2", "Gene", "InnerPrimerSeq"];
opts.VariableTypes = ["string", "string", "string", "string"];
opts = setvaropts(opts, ["pEXCFbackbone_V1", "V2", "Gene", "InnerPrimerSeq"], "WhitespaceRule", "preserve");
opts = setvaropts(opts, ["pEXCFbackbone_V1", "V2", "Gene", "InnerPrimerSeq"], "EmptyFieldRule", "auto");
% Read the data from the specified Excel file into a table
TFI_extraction_Seq = readtable(TFI_sheet, opts, "UseExcel", false);
clear opts

% Extract relevant columns from the table as strings
V1 = string(TFI_extraction_Seq{1,1});
V2 = string(TFI_extraction_Seq{:,2});
TFI_gene_name = string(TFI_extraction_Seq{:,3});
R1_primer = string(TFI_extraction_Seq{:,4});

% Create a new table with the extracted gene information
TFI_Gene_table = table([TFI_gene_name,V2,R1_primer]);
TFI_Gene_table = splitvars(TFI_Gene_table);
TFI_Gene_table.Properties.VariableNames = {'Gene ID','V2','Inner Primer'};

% Initialize a count array for TFI
Count_TFI = zeros(length(R1_primer),1);

% Read the sequences from the FASTQ file
read_sequ= fastqread(filename_fastq);
temp = struct2cell(read_sequ.').'; 
[m1,~] = size(temp);
Sequ = temp(1:m1,2);
[tot_seq, ~]  = size(temp);

% Count occurrences of primers in the sequences
for n1= 1:length(R1_primer)
    % Iterate over each sequences in fastq file
    for k=1:m1
        Sequ1= string(Sequ(k));
        % Check if Sequ1 contains both R1_primer and V2 for the current indices
        TF(k) = (contains(Sequ1,R1_primer(n1))&contains(Sequ1,V2(n1)));
    end
    TF = double(TF); % Convert logical array to double
    % Count the number of true values in TF for the current n1
    Count_TFI(n1) = sum(TF);
end
% Convert the count array to a table
Count_TFI = table(Count_TFI);
Count_TFI.Properties.VariableNames = {Count_TFI};

% Combine the gene table with the counts and write to an Excel file
TFI_Gene_table = [TFI_Gene_table,Count_TFI];
writetable(TFI_Gene_table,Output_filename,'Sheet','Count_TFI');
