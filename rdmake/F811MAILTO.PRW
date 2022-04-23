#Include 'Protheus.ch'

user function F811MAILTO()

//local cEmail := "teste1@totvs.com;teste2@totvs.com;teste3@totvs.com"
local cRet := ''

if !empty(SA1->A1_XEMCOB)
      cRet := SA1->A1_XEMCOB
endIf

return cRet
