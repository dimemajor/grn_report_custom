# -*- coding: utf-8 -*-
{
    'name': "Custom Goods Receipt Report",

    'summary': """
        Goods Recipt Report within a period | Goods Receipt of a vendor | Receipt Report of a product within a period, with cost details """,

    'description': """
        Shows expiration date and vendor reference in the Goods Recieved Report. 
        Depends on GRN written by Loyal IT Solutions Pvt Ltd
    """,

    'author': "Atsevah Anthony",
    'website': "",

    'category': 'uncategorized',
    'version': '15.0.0',
    'license': 'AGPL-3',

    'depends': ['grn_report', 'product_expiry'],

    'data': [
        'views/views.xml',
    ],

}
