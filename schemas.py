class SmallSchema():
    id = "id"
    fname = "fname"
    infix = "infix"
    lname = "lname"


class CommitteeExcel(SmallSchema):
    total = "total"
    bedrag_a = "bedrag_a"
    bedrag_b = "bedrag_b"
    bedrag_c = "bedrag_c"


class BigSchema(SmallSchema):
    email = "email"
    iban = "iban"
    countrycode = "countrycode"


class ClubCollectSchema(BigSchema):
    phone = "phone"


# class ExcelSchema(BigSchema):
#     roepnaam: Series[str]
#     bic: Series[str]
#     iban_type: Series[str]
#     sepa_date_of_signature: Series[str]
#     address: Series[str]
#     postal_code: Series[str]
#     city: Series[str]
