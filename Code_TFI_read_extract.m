clear 
clc
warning( 'off', 'MATLAB:xlswrite:AddSheet' ) ;
Range_1 = [" A1"; " B1"; " C1"; " D1"; " E1"; " F1"; " G1"; " H1"; " I1"; " J1"; " K1";  " L1"; " M1"; " N1"; " O1"; " P1"; " Q1"; " R1"; " S1";  " T1"; " U1"; " V1"; " W1";  " X1"; " Y1"; " Z1";...
"AA1"; "AB1"; "AC1"; "AD1"; "AE1"; "AF1"; "AG1"; "AH1"; "AI1"; "AJ1"; "AK1";  "AL1"; "AM1"; "AN1"; "AO1"; "AP1"; "AQ1"; "AR1"; "AS1";  "AT1"; "AU1"; "AV1"; "AW1";  "AX1"; "AY1"; "AZ1";...
"BA1"; "BB1"; "BC1"; "BD1"; "BE1"; "BF1"; "BG1"; "BH1"; "BI1"; "BJ1"; "BK1";  "BL1"; "BM1"; "BN1"; "BO1"; "BP1"; "BQ1"; "BR1"; "BS1";  "BT1"; "BU1"; "BV1"; "BW1";  "BX1"; "BY1"; "BZ1";...
"CA1"; "CB1"; "CC1"; "CD1"; "CE1"; "CF1"; "CG1"; "CH1"; "CI1"; "CJ1"; "CK1";  "CL1"; "CM1"; "CN1"; "CO1"; "CP1"; "CQ1"; "CR1"; "CS1";  "CT1"; "CU1"; "CV1"; "CW1";  "CX1"; "CY1"; "CZ1";...
"DA1"; "DB1"; "DC1"; "DD1"; "DE1"; "DF1"; "DG1"; "DH1"; "DI1"; "DJ1"; "DK1";  "DL1"; "DM1"; "DN1"; "DO1"; "DP1"; "DQ1"; "DR1"; "DS1";  "DT1"; "DU1"; "DV1"; "DW1";  "DX1"; "DY1"; "DZ1";...
"EA1"; "EB1"; "EC1"; "ED1"; "EE1"; "EF1"; "EG1"; "EH1"; "EI1"; "EJ1"; "EK1";  "EL1"; "EM1"; "EN1"; "EO1"; "EP1"; "EQ1"; "ER1"; "ES1";  "ET1"; "EU1"; "EV1"; "EW1";  "EX1"; "EY1"; "EZ1";...
"FA1"; "FB1"; "FC1"; "FD1"; "FE1"; "FF1"; "FG1"; "FH1"; "FI1"; "FJ1"; "FK1";  "FL1"; "FM1"; "FN1"; "FO1"; "FP1"; "FQ1"; "FR1"; "FS1";  "FT1"; "FU1"; "FV1"; "FW1";  "FX1"; "FY1"; "FZ1";...
"GA1"; "GB1"; "GC1"; "GD1"; "GE1"; "GF1"; "GG1"; "GH1"; "GI1"; "GJ1"; "GK1";  "GL1"; "GM1"; "GN1"; "GO1"; "GP1"; "GQ1"; "GR1"; "GS1";  "GT1"; "GU1"; "GV1"; "GW1";  "GX1"; "GY1"; "GZ1";...
"HA1"; "HB1"; "HC1"; "HD1"; "HE1"; "HF1"; "HG1"; "HH1"; "HI1"; "HJ1"; "HK1";  "HL1"; "HM1"; "HN1"; "HO1"; "HP1"; "HQ1"; "HR1"; "HS1";  "HT1"; "HU1"; "HV1"; "HW1";  "HX1"; "HY1"; "HZ1";...
"IA1"; "IB1"; "IC1"; "ID1"; "IE1"; "IF1"; "IG1"; "IH1"; "II1"; "IJ1"; "IK1";  "IL1"; "IM1"; "IN1"; "IO1"; "IP1"; "IQ1"; "IR1"; "IS1";  "IT1"; "IU1"; "IV1"; "IW1";  "IX1"; "IY1"; "IZ1";...
"JA1"; "JB1"; "JC1"; "JD1"; "JE1"; "JF1"; "JG1"; "JH1"; "JI1"; "JJ1"; "JK1";  "JL1"; "JM1"; "JN1"; "JO1"; "JP1"; "JQ1"; "JR1"; "JS1";  "JT1"; "JU1"; "JV1"; "JW1";  "JX1"; "JY1"; "JZ1";...
"KA1"; "KB1"; "KC1"; "KD1"; "KE1"; "KF1"; "KG1"; "KH1"; "KI1"; "KJ1"; "KK1";  "KL1"; "KM1"; "KN1"; "KO1"; "KP1"; "KQ1"; "KR1"; "KS1";  "KT1"; "KU1"; "KV1"; "KW1";  "KX1"; "KY1"; "KZ1";...
"LA1"; "LB1"; "LC1"; "LD1"; "LE1"; "LF1"; "LG1"; "LH1"; "LI1"; "LJ1"; "LK1";  "LL1"; "LM1"; "LN1"; "LO1"; "LP1"; "LQ1"; "LR1"; "LS1";  "LT1"; "LU1"; "LV1"; "LW1";  "LX1"; "LY1"; "LZ1"];
warning('off')
[~, ~, raw] = xlsread('G:\Sequence analysis_Jees_BARCODE_EXPERIMENTS_2020\TFOE_2020\TFOE  REPEAT_10XRIF NOV9-2020\TFI V1 AND V2 SEQ FOR EXTRACTION.xlsx','Sheet1');
% V1 - pEXCFbackbone, V2 - TF specific sequence

stringVectors = string(raw(3:end,[1,2,3,4]));
stringVectors(ismissing(stringVectors)) = '';
V1  = stringVectors(1,1);
R1_V2  = stringVectors(:,2);
R1_primer  = stringVectors(:,4);
lenght_V2 = length(R1_V2);
for k=1:length(R1_V2)
    R1_V2_k = R1_V2(k,1);
    R1_V2_k = convertStringsToChars(R1_V2_k);
    R1_V2_k_com1 = seqrcomplement(R1_V2_k);
    R1_V2_k_com1 = strcat(R1_V2_k_com1);
    R2_V2 (k,:)= cellstr(R1_V2_k_com1);
   
    R2_primer_k =R1_primer(k,1);
    R2_primer_k = convertStringsToChars(R2_primer_k);
    R2_primer_k_com1 = seqrcomplement(R2_primer_k);
    R2_primer_k_com1 = strcat(R2_primer_k_com1);
    R2_primer (k,:)= cellstr(R2_primer_k_com1);
end
 R2_V2 = string(R2_V2);
  R1_V2 = string(R1_V2);
filename_xlsx = 'Code_1_TFI_COUNT.xlsx';
Count_TFI_R11 = table(R1_V2);
Count_TFI_R22 = table(R2_V2);
Count_TFI_R11.Properties.VariableNames = {'TFI_R1_V2'};
Count_TFI_R22.Properties.VariableNames = {'TFI_R2_V2'};
writetable(Count_TFI_R11,filename_xlsx,'Sheet','TFI_R1','Range',char(Range_1(1)))
writetable(Count_TFI_R22,filename_xlsx,'Sheet','TFI_R2','Range',char(Range_1(1)))


cutoff = 10;
for i = 1:10 % fastq_q file number
    Count_TFI_R1 = zeros(length(R1_V2),1);
    Count_TFI_R2 = zeros(length(R1_V2),1);
    tic
    disp(i)
for j=1:2 % j = 1 for R1 and j = 2 for R2
S1 = 'Fastq_files_TFI\';
S2 = num2str(i);
S3 = '_S';
if j==1
    S4='_L001_R1_001.fastq';
else
    S4='_L001_R2_001.fastq';
end
filename_fastq = strcat(S1,S2,S3,S2,S4);
sheet_name = strcat(S2,S3,S2);
read_sequ= fastqread(filename_fastq);
temp = struct2cell(read_sequ.').'; 
[m1,~] = size(temp);
Sequ = temp(1:m1,2);
[tot_seq, ~]  = size(temp);
x = 1;
if j ==1
    for n1= 1:length(R1_V2)
        for k=1:m1
            Sequ1= string(Sequ(k));
            TF(k) = (contains(Sequ1,R1_primer(n1))&contains(Sequ1,R1_V2(n1)));
        end
        TF = double(TF);
        Count_TFI_R1(n1) = sum(TF);
    end
else
    for n1= 1:length(R1_V2)
        for k=1:m1
            Sequ1= string(Sequ(k));
            RTF(k) = (contains(Sequ1,R2_primer(n1))&contains(Sequ1,R2_V2(n1)));
            TF(k) = (contains(Sequ1,R2_V2(n1)));
        end
        TF = double(TF);
        Count_TFI_R2(n1) = sum(TF);
    end
end
  
end

V1 = 'TFI_S';
V2 = num2str(i);
V3 = '_R1';
V4 = '_R2';
R1_name = strcat(V1,V2,V3);
R2_name = strcat(V1,V2,V4);
Count_TFI_R1_1 = table(Count_TFI_R1);
Count_TFI_R2_1 = table(Count_TFI_R2);
Count_TFI_R1_1.Properties.VariableNames = {R1_name};
Count_TFI_R2_1.Properties.VariableNames = {R2_name};
writetable(Count_TFI_R1_1,filename_xlsx,'Sheet','TFI_R1','Range',char(Range_1(i+1)))
writetable(Count_TFI_R2_1,filename_xlsx,'Sheet','TFI_R2','Range',char(Range_1(i+1)))
toc
end