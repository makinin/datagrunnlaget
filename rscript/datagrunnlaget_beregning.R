## Laster pakker ----
library(rdbhapi)
library(tidyverse)
library(reader)
library(readxl)
library(openxlsx)


# DBH-tabeller for Blåtthefte med API-valg


finsystabeller <-
  tribble(~tabellnavn, ~table_id, ~group_by,~filters,
          "institusjoner", 211, NULL, NULL,
          "institusjonstyper",287,NULL,NULL,
          "studiepoeng", 900, NULL, list("Årstall"=c("top",2)),
          "doktorgrader", 101, c("Institusjonskode", "Årstall"), list("Årstall"=c("top",2)),
          "doktorgrader_samarbeid", 100, c("Årstall", "Institusjonskode (arbeidsgiver)"), list("Årstall"=c("top",2)),
          "PKU", 98, NULL, list("Årstall"=c("top",2)),
          "kandidater", 907, NULL, list("Årstall"=c("top",2)),
          "publisering", 373, c("Årstall", "Institusjonskode"), list("Årstall"=c("top",2)),
          "utveksling", 142, c("Årstall", "Institusjonskode", "Utvekslingsavtale","Type", "Nivåkode"), list("Årstall"=c("top",2)),
          "økonomi", 902, NULL, list("Årstall"=c("top",2))
  )



#' Hjelpefunksjon som laster ned finsystabellene
#'
#' @param tabellnavn 
#' @param table_id 
#' @param group_by 
#' @param filters 
#'
#' @return
#' @export
#'
#' @examples
hent_dbh_data_wrapper <- function(tabellnavn, table_id, group_by,  filters) {
  pb$tick()$print()
  res <- 
    do.call(dbh_data,
            c(list(table_id = table_id,
                   group_by = group_by, filters=filters)
            ))%>%
    
    rename_all(str_to_lower)
  assign(str_c(tabellnavn, "_org"), res, envir = .GlobalEnv)
  return(tabellnavn)
}

# Laster ned alle finsystabellene

pb <- progress_estimated(nrow(finsystabeller))
pmap(finsystabeller, hent_dbh_data_wrapper)


## Håndterer institusjonsendringer ----


institusjoner <-
  institusjoner_org %>%
  mutate(institusjonskode_nyeste = institusjoner_org$`institusjonskode (sammenslått)`, institusjonskode_gammel=institusjoner_org$institusjonskode
           ) %>% merge(institusjonstyper_org, by="institusjonstypekode", all.x = TRUE)

## Leser manuelle tabeller for satser mm. ---

c("tilskuddsgrad",
  "unntak",
  "institusjoner_finsys") %>% 
  walk(~assign(.,
               as_tibble(read_xlsx("data/finsysinfo.xlsx", sheet = .)),
               
               envir = .GlobalEnv))


unntak <- unntak %>% 
  complete(budsjettår = 2020:2021,
           nesting(indikator, institusjonskode))

# Inkluderer også tall fra forløperinstitusjoner til de som er inkludert for 2020.
institusjoner_finsys <- institusjoner_finsys %>% 
  bind_rows(semi_join(institusjoner,
                      institusjoner_finsys,
                      by = c("institusjonskode_nyeste" = "institusjonskode")) %>% 
              select(institusjonskode = institusjonskode_gammel)) %>% 
  complete(budsjettår = 2020:2021,
           institusjonskode) %>% 
  drop_na(budsjettår)


## Filtrerer finsystabeller til hva som gir uttelling, og harmoniserer til sammenslåing ----

studiepoeng <-
  studiepoeng_org %>%
  rename(indikatorverdi = `ny produksjon egentfin`) %>%
  # Bruker finansieringskategorien basert på studentens tilhørighet for BI og for emnet ellers
  mutate(kategori =
           case_when(institusjonskode == "8241" ~
                       `finmodkode student`,
                     TRUE ~
                       `finmodekode emne`)) %>% 
  mutate_at(c("studentkategori", "kategori"), str_to_upper) %>% 
  filter(kategori %in% LETTERS[1:6],
         studentkategori == "S")

kandidater <-
  kandidater_org %>% 
  rename(faktor = `uttelling kode:1=enkel, 2=dobbel`,
         etterrapportert = etterrapp,
         innpasset = `dobbel til enkel`,
         kategori = `finansierings-kategori`) %>% 
  mutate_at("kategori", str_to_upper) %>%
  mutate_at("faktor", parse_integer) %>%
  mutate_at(c("etterrapportert", "innpasset"), replace_na, 0) %>% 
  mutate(ordinær = totalt - etterrapportert - innpasset) %>% 
  select(-totalt) %>% 
  gather(ordinær, etterrapportert, innpasset,
         key = "kandidatgruppe",
         value = "indikatorverdi")

erasmus_plus<-utveksling_org %>% 
  rename(indikatorverdi = `antall totalt`) %>% 
  mutate_at(c("utvekslingsavtale", "type"), str_to_upper) %>%
  filter(utvekslingsavtale == "ERASMUS+" & nivåkode!= "FU" & type=="NORSK") %>% 
  mutate(kategori="Erasmus+")

utenlandsk<- utveksling_org %>% 
  rename(indikatorverdi = `antall totalt`) %>% 
  mutate_at(c("utvekslingsavtale", "type"), str_to_upper) %>%
  filter(utvekslingsavtale != "INDIVID" & nivåkode!= "FU" & type=="UTENL") %>% 
  mutate(kategori="ordinær")

utreisende_norsk<- utveksling_org %>% 
  rename(indikatorverdi = `antall totalt`) %>% 
  mutate_at(c("utvekslingsavtale", "type"), str_to_upper) %>%
  filter(!utvekslingsavtale %in% c("ERASMUS+","INDIVID") & nivåkode!= "FU" & type=="NORSK") %>% 
  mutate(kategori="ordinær")

utveksling<-bind_rows(erasmus_plus,utenlandsk,utreisende_norsk)

doktorgrader <- 
  doktorgrader_org %>% 
  rename(indikatorverdi = `antall totalt`) %>%
  add_column(kandidatgruppe = "ordinær") %>% 
  bind_rows(
    doktorgrader_samarbeid_org %>%
      rename(institusjonskode = `institusjonskode (arbeidsgiver)`) %>% 
      count(årstall, institusjonskode, name = "indikatorverdi") %>% 
      add_column(kandidatgruppe = "samarbeids-ph.d.",
                 faktor = 0.2),
    PKU_org %>% 
      rename(indikatorverdi = antall) %>% 
      add_column(kandidatgruppe = "PKU"))

publisering <- 
  publisering_org %>% 
  rename(indikatorverdi = publiseringspoeng)

økonomi <- 
  økonomi_org %>% 
  rename(EU = eu) %>% 
  mutate(forskningsråd = map2_dbl(nfr, rff, sum, na.rm = TRUE),
         BOA = map2_dbl(bidrag, oppdrag, sum, na.rm = TRUE))

c("EU", "forskningsråd", "BOA") %>% 
  walk(~assign(.,
               select_at(økonomi, c("institusjonskode", "årstall", "indikatorverdi" = .)),
               envir = .GlobalEnv))

## Lager samletabell og beregner endringer og uttelling ----

finsys_data <-
  # Samler alle enkelttabellene
  c("studiepoeng", "kandidater", "utveksling", "doktorgrader", "publisering", "EU", "forskningsråd", "BOA") %>%
  setNames(., .) %>% 
  map_df(get, .id = "indikator") %>%
  mutate(budsjettår = årstall + 2) %>% 
  # Summerer for variablene som skal beholdes
  group_by(
    budsjettår,
    institusjonskode,
    indikator, 
    kategori, 
    kandidatgruppe,
    faktor) %>%
  summarise_at("indikatorverdi", sum, na.rm = TRUE) %>%
  ungroup %>% 
  # Fyller ut med 0 for manglende kombinasjoner (forskring for å få endringstall riktig)
  complete(budsjettår = full_seq(budsjettår, 1),
           institusjonskode,
           nesting(indikator, 
                   kategori, 
                   kandidatgruppe,
                   faktor),
           fill = list(indikatorverdi = 0)) %>%
  # Legger til endringstall fra året før
  group_by(
    institusjonskode,
    indikator, 
    kategori, 
    kandidatgruppe,
    faktor) %>% 
  arrange(budsjettår) %>% 
  mutate(indikatorendring = indikatorverdi - lag(indikatorverdi)) %>%
  ungroup %>% 
  # Filtrerer ut institusjoner som ikke inngår i systemet eller har unntak for indikatorer
  semi_join(institusjoner_finsys,
            by = c("institusjonskode", "budsjettår")) %>% 
  anti_join(unntak,
            by = c("institusjonskode", "budsjettår", "indikator")) %>%
  
  left_join(tilskuddsgrad,
            by = c("institusjonskode", "budsjettår", "indikator")) %>% 
  
  # Legger til institusjonsinfo
  left_join(select(institusjoner,
                   institusjonskode, institusjonsnavn, kortnavn, institusjonskode_nyeste),
            by = "institusjonskode") %>% 
  left_join(select(institusjoner_finsys,
                   institusjonskode,  budsjettår),
            by = c("institusjonskode","budsjettår")) %>%  
  mutate(budsjettår=as.factor(budsjettår))



## Lager tabeller med produksjonsdata, tilsvarende dem i Blått hefte ----

#' Title
#'
#' @param df 
#' @param filter_indikatorer 
#' @param kolonne_vars 
#' @param funs 
#'
#' @return
#' @export
#'
#' @examples
lag_produksjonstabell <- function(df,
                                  filter_indikatorer, # indikatorene som skal inngå i tabellen
                                  kolonne_vars, # variablene som skal spres på kolonner (indikator/kategori/osv.)
                                  funs # tilleggsfunksjon som skal anvendes på kolonnene (for å lage andelstall)
) {
  if (is.na(funs)) {
    funs <- identity
  }
  df %>% 
    filter(indikator %in% filter_indikatorer) %>% 
    unite_("kolonne_vars_samlet", kolonne_vars) %>% 
    group_by_at(c("institusjonskode_nyeste",
                  "institusjonsnavn",
                  "kortnavn",
                  "kolonne_vars_samlet",
                  "budsjettår")) %>% 
    summarise_at("indikatorverdi", sum) %>% 
    ungroup() %>% 
    spread(key = "kolonne_vars_samlet", value = "indikatorverdi") %>%
    mutate_if(is.numeric, funs)
}


produksjonstabeller <- 
  tribble(~filter_indikatorer, ~kolonne_vars, ~funs,
          "studiepoeng", "kategori", NA,
          "kandidater", c("kategori", "faktor"), NA,
          c("doktorgrader", "utveksling"), c("indikator", "faktor", "kategori"), NA,
          c("publisering", "EU", "forskningsråd", "BOA"),c("indikator") , list(pst = function(x) 100 * x / sum(x, na.rm = TRUE))) %>% 
  pmap(function(...) {
    finsys_data %>% 
      filter(budsjettår == 2021) %>% 
      lag_produksjonstabell(...)
  })

# expot i excel 


studiepoeng<-produksjonstabeller[[1]]
kandidater<-produksjonstabeller[[2]]
utveksling_doktorgrader<-produksjonstabeller[[3]]
publisering_eu_nfr_boa<-produksjonstabeller[[4]]

wb <- createWorkbook()
addWorksheet(wb, "studiepoeng")
addWorksheet(wb, "kandidater")
addWorksheet(wb, "utveksling_doktorgrader")
addWorksheet(wb, "publisering_eu_nfr_boa")

writeDataTable(wb, "studiepoeng", x = studiepoeng)
writeDataTable(wb, "kandidater", x = kandidater)
writeDataTable(wb, "utveksling_doktorgrader", x = utveksling_doktorgrader)
writeDataTable(wb, "publisering_eu_nfr_boa", x = publisering_eu_nfr_boa)

saveWorkbook(wb, "Blåttheftetabeller.xlsx", overwrite = TRUE)
