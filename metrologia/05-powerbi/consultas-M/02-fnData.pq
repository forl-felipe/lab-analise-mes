// Função auxiliar: converte data do SharePoint (que pode vir com fuso) em data simples, sem deslocar o dia.
// Nome da consulta: fnData
(v as any) as nullable date =>
    if v = null then null
    else if Value.Is(v, type datetimezone) then Date.From(DateTimeZone.RemoveZone(v))
    else if Value.Is(v, type datetime) then Date.From(v)
    else if Value.Is(v, type date) then v
    else Date.From(DateTime.From(v, "en-US"))
