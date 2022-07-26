# ODK2Doc 📱➡️📝 
Aid organisations, academic researchers and civil society organisations often rely on data collection platforms such as [ODK/Kobo](https://getodk.org/vs-kobo/). These survey applications use a common format for designing questionnaires, which are then used by teams and enumerators to collect data via an offline mobile phone application.

People working with these platforms, at times, need to convert their questionnaires to [printable formats](https://forum.getodk.org/t/download-form-to-word/5868) or [word documents](https://community.kobotoolbox.org/t/do-we-convert-kobo-question-form-into-microsoft-word/5177). This app tries to make this process easier and more straightforward. This app converts the xls survey form (Used by Kobo, ODK & even, SurveyCTO - although, SurveyCTO allows users to download html versions of their forms, for printing!) to formatted word documents for printing and/or adding to reports. 

The final app is available for use [here](https://zaeendesouza.shinyapps.io/ODK2Doc/). I am still testing it out on various different forms, but if, while using it, you find any issues or errors, please flag these. Also, if you would like a specific feature(s) to be added, please let me know and I will try my best to add them!


**Note:** *This is a recent version of this app; I do plan to add more features soon.*

## Tips

Make sure that the form has a title, that has been added via the 'settings' sheet of the form. The app also requires that sheets have the default names:

1. survey
2. choices (Within the choice sheet, you should choose one label column which you want to use and name is "label" as opposed to "label::english" or "label::[your language]")
3. settings

Other names, will not work with this app.

## Files

1. **App:** *Main shiny app file (contains ui and server).*
2.  **report1_compact_skip.Rmd:** *Template for when people choose to add skip logic.*
3. **report2_compact.Rmd:** *Template for minimal form.*
4. **styles.css:** *Style sheet.*
