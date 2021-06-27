Public NotInheritable Class AboutOrklaRT

    Private Sub AboutOrklaRT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ApplicationTitle As String
        If Not String.IsNullOrWhiteSpace(My.Application.Info.Title) Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

        Me.LabelProductName.Text = "Orkla Reporting Tool"
        Me.LabelVersion.Text = String.Format("Version {0}", "4.2")
        Me.LabelCopyright.Text = "Copyright © Orkla Foods Norge 2018"
        Me.LabelCompanyName.Text = "developed by OrklaIT AS"
        Me.TextBoxDescription.Text = "Release Note v4.2:" & vbNewLine &
                                    "Feil rettet i produksjon plan rapport sequence og Allergen  " & vbNewLine &
                                    "SAP Factory Calender er oppdater til 2025 " & vbNewLine &
                                    "Release Note v4.1:" & vbNewLine &
                                    "Endring i produksjon plan rapport og blande plan med AllergenType informasjon " & vbNewLine &
                                    "Release Note v4.0:" & vbNewLine &
                                    "Feil rettet i LagerHistorikk rapport " & vbNewLine &
                                    "Release Note v3.9:" & vbNewLine &
                                    "Feil rettet ordre nummer viser tekst " & vbNewLine &
                                    "Status produksjon,Prosess og Produksjonsordre historikk " & vbNewLine &
                                    "Daglig produksjons plan, MRP Behov " & vbNewLine &
                                    "Release Note v3.8:" & vbNewLine &
                                    "1. Produksjonsplan: Feilretting ordre nummer i rapport " & vbNewLine &
                                    "2. Blandeplan: Feilretting ordre nummer i rapport " & vbNewLine &
                                    "Release Note v3.7:" & vbNewLine &
                                    "1. Produksjonsplan: Denne skal respondere raskere i ny versjon " & vbNewLine &
                                    "2. Blandeplan : Høyre klikk rett inn i SAP fra ORT skal nå fungere " & vbNewLine &
                                    "   Allergen skal oppdateres første gang rapport kjøres. " & vbNewLine &
                                    "3. Holdbarhetsrapport: Kommentarfelt skal nå lagres " & vbNewLine &
                                    "   Materialplanlegger er med i utvalgsmeny og skal dermed gi for alle fabrikker " & vbNewLine &
                                    "   Beholdninger som ikke finnes er fjernet " & vbNewLine &
                                    "4. Daglig salgsordre: Ikke_Levert- kollonne; Kvantum blankt når levert kvantum 0 " & vbNewLine &
                                    "   Mankooversikt inkluderer problemvarer med årsakkode i utvalg rapport " & vbNewLine &
                                    "5. Kapasitetsbilde pr ressurs: Stranda fikk ikke ut data. " & vbNewLine &
                                    "6. Bristoversikt: Utvalg for rapport muliggjør at fabrikk er valgfri, men materialtype pliktig " & vbNewLine &
                                    "7. Daglig produksjon: Manglende produktgruppe oppdatert " & vbNewLine &
                                    "   Tilpasset rapport forhåndsdefinert for Elverum – tilleggsfelter og layout " & vbNewLine &
                                    "8. Lager verdi og dekning dager: Mulig å velge Materialart for rapportutvalg (1 eller flere) " & vbNewLine &
                                    "Release Note v3.6:" & vbNewLine &
                                    "1. OrklaRT v3.6 endringene " & vbNewLine &
                                    "   ny felt Materialstatus i Bristoversikt, Holdbarhetrapport og Salgsordrestatus" & vbNewLine &
                                    "   ny felt Allergen i Produksjonplan og Blandeplan" & vbNewLine &
                                    "2. Ny rapport Where Used MultiLevel." & vbNewLine &
                                    "Release Note v3.5:" & vbNewLine &
                                    "1. Holdbarhetsrapport: feltene Produktansvarlig og Divisjon til utvalgslisten, ny felt viser Verdi utgått batch, " & vbNewLine &
                                    "   viser hvor mange paller (TPK) rest på batch tilsvarervalgfri fabrikk felt." & vbNewLine &
                                    "2. Viser antall DPAK i daglig produksjonsplan rapport." & vbNewLine &
                                    "Release Note v3.4:" & vbNewLine &
                                    "1. Kondemneringsoversikt: valgfri fabrikk felt." & vbNewLine &
                                    "2. Lagring av pivot layouter." & vbNewLine &
                                    "3. Ny skjerm(Bruker Oppsett) til operette ny bruker i OrklaRT." & vbNewLine &
                                    "4. MD04: Nå fungerer for alle fabrikker." & vbNewLine &
                                    "5. Høyreklikk på material for å ser på inngår i eller består av." & vbNewLine &
                                    "6. Feil rettet i Lager Historikk og Beholdnings simulering." & vbNewLine &
                                    "7. Holdbarhet: Restbeholdning beregnes fram i tid basert på prognose for å kunne identifisere mulig feilkost for tidlig nok å gjøre tiltak."
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub

End Class
