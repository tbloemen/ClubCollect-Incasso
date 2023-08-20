# ClubCollect Incasso

The following graph describes the flow of the incasso.

```mermaid
flowchart TD
Start --> B
B[read member data] --Determine type of incasso--> E["Parse member \ninformation and put \ninto dataframe"]

subgraph Main Program
direction TB
subgraph Handle cases polymorphicly
E --> G["Combine to make \na filled dataframe"]
G --> H(["Put dataframe with \ncorrect formatting \nin Excel"])
end

E ==> Create([Create a new \nblanc small list \nfor committees])

E -..-> F
F[Combine \nsmall lists \non id]--> G
H --> Move([Move processed \nsmall lists to \nthe processed folder])
Create

end
H --Do by mail merge---> I1[Send Emails]
H --Can't be done \nprogrammatically---> I2

Start ==> J1

subgraph Access/Excel
I1 --After 7 days--> J1([Fill in Recurr file])
J1 -- Can't be done \nprogrammatically--> K1[Import Recurr into IBANC]
K1 --> L1[Upload IBANC xml to ING]
end

subgraph ClubCollect
I2["Paste information into \nclubcollect site"]
end

```

After the main program, manual action is needed, regardless of which type of member administration is used. Except for filling the Recurr file, that is something the program can do (currently not implemented). However, that option is only possible after determining that the member admin type is of the Access/Excel kind.
