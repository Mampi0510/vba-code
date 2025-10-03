Function BORNE_INFERIEURE (plage as Range) as double
    dim cell as plage
    dim bi as double

    for  Each cell in plage
        if cell.Value > bi Then
        bi= cell.Value
        end if
    next cell

    BORNE_INFERIEURE= bi
End Function