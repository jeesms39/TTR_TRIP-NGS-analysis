
%RIF-FITC SORT FRACTIONS OF KNOCKDOWN STRAINS
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

[~, ~, raw] = xlsread('MiSeq_extraction_sequences.xlsx','Sheet1');
EXP3 = string(raw(2:20,1:2));
EXP3_gene_name = EXP3(1:end-2,1);
EXP3_Prim = EXP3(1:end-2,2);
V1 = EXP3(end-1,2);
V2 = EXP3(end,2);
EXP3_T =  [table(EXP3_gene_name),table(EXP3_Prim)];
EXP3_T.Properties.VariableNames = {'Name', 'Gene'};
filename_xlsx = 'Code_1_MiSeq_sequences_COUNT.xlsx';
writetable(EXP3_T,filename_xlsx,'Sheet','MiSeq','Range',char(Range_1(1)))
cutoff = 10;
clearvars raw EXP3 EXP3_gene_name EXP3_T

for i = 68 % fastq_q file number
    tic
    disp(i)
for j=1 % j = 1 for R1 and j = 2 for R2
S1 = 'MiSeq_extraction\';
S2 = 'S';
S3 = num2str(i);
if j==1
    S4='_L001_R1_001.fastq';
else
    S4='_L001_R2_001.fastq';
end
S6 ='.xlsx';
filename_fastq = strcat(S1,S2,S3,S4);
clearvars S1 S2 S3 S4 S6
read_sequ= fastqread(filename_fastq);
temp = struct2cell(read_sequ.').'; 
[m1,~] = size(temp);
Sequ = temp(1:m1,2);
clearvars read_sequ temp
    for n1= 1:length(EXP3_Prim)
        for k=1:m1
            Sequ1= string(Sequ(k));
            TF1(k) = (contains(Sequ1,EXP3_Prim(n1))&contains(Sequ1,V1)&contains(Sequ1,V2));
            TF2(k) = (contains(Sequ1,EXP3_Prim(n1)));
        end
        TF1 = double(TF1);
        Count1(n1,1) = sum(TF1);
        TF2 = double(TF2);
        Count2(n1,1) = sum(TF2);
        clearvars TF1 TF2
    end
end

A1 = 'S';
A2 = num2str(i);
R1_name = strcat(A1,A2);
Count_R1 = table(Count1);
Count_R1.Properties.VariableNames = {R1_name};
writetable(Count_R1,filename_xlsx,'Sheet','MiSeqV1V2','Range',char(Range_1(i+2)))
Count_R2 = table(Count2);
Count_R2.Properties.VariableNames = {R1_name};
writetable(Count_R2,filename_xlsx,'Sheet','MiSeq','Range',char(Range_1(i+2)))
clearvars Count1 Count2  temp m1 filename_fastq Sequ TF1 TF2 m1 n1 filename_fastq
end
clearvars i

clearvars -except Range_1

