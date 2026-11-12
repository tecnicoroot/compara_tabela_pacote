select 
c.CD_CONVENIO,
c.ds_convenio Convenio,
obter_codigo_pacote_acomod(b.nr_seq_pacote) codigos ,

obter_valor_dominio(56, b.ie_tipo_acomod) Classif_acomodacao,
obter_categoria_convenio(obter_convenio_pacote(b.nr_seq_pacote), b.cd_categoria) categoria,
b.nr_sequencia,
b.DT_VIGENCIA,
b.DT_VIGENCIA_FINAL,
b.cd_procedimento,
a.cd_proced_pacote,
obter_valor_dominio(30, a.ie_origem_proced) Tabela,
b.VL_PACOTE valor_pacote_tasy


from PACOTE a 
inner join PACOTE_TIPO_ACOMODACAO b on b.nr_seq_pacote = a.nr_seq_pacote
inner join convenio c on c.cd_convenio = a.cd_convenio
where 1=1
and a.ie_situacao = 'A' -- pacotes ativos
and c.IE_SITUACAO = 'A' -- Acomodacao
and c.cd_convenio in (3,5) -- Juiz de Fora e Intercambio
--and b.DT_VIGENCIA_FINAL between  to_date('01/07/2026 00:00:00', 'dd/mm/yyyy hh24:mi:ss') and  to_date('31/12/2050 23:59:59', 'dd/mm/yyyy hh24:mi:ss') --com data de vigência  maior que a data atual
--and a.cd_proced_pacote = 80808043 --exemplo
and b.CD_PROCEDIMENTO in (
80808045,

80808067

)