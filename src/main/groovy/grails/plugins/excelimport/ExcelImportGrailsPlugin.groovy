package grails.plugins.excelimport

import grails.plugins.*

class ExcelImportGrailsPlugin extends Plugin {
    // the plugin version
    def version = "2.0.0-SNAPSHOT"
    // the version or versions of Grails the plugin is designed for
    def grailsVersion = "3.0.1 > *"
    // the other plugins this plugin depends on
    def dependsOn = ["joda-time":"2.0 > *"]
    // resources that are excluded from plugin packaging
    def pluginExcludes = [
            "grails-app/views/error.gsp"
    ]

    // TODO Fill in these fields
    def author = "Jean Barmash, Oleksiy Symonenko"

    def authorEmail = "Jean.Barmash@gmail.com"
    def title = "Excel, Excel 2007 & CSV Importer Using Apache POI"
    def description = '''\\
	Excel-Import plugin uses Apache POI [http://poi.apache.org/] library (v 3.6) to parse Excel files.  
      	It's useful for either bootstrapping data, or when you want to allow your users to enter some data using Excel spreadsheets. 
'''
    def profiles = ['web']

	def license = "APACHE"
	def organization = [ name: "EnergyScoreCards.com", url: "http://www.energyscorecards.com/" ]
	def developers = [
        	[ name: "Oleksiy Symonenko", email: "" ],
        	[ name: "Puneet Behl", email: "puneet.behl007@gmail.com" ]
	]

	def issueManagement = [ system: "JIRA", url: "http://jira.grails.org/browse/GPXLIMPORT" ]
	def scm = [ url: "https://github.com/jbarmash/grails-excel-import" ]
    // URL to the plugin's documentation
    def documentation = "http://grails.org/plugin/excel-import"

    Closure doWithSpring() { {->
        // TODO Implement runtime spring config (optional)
    } }

    void doWithDynamicMethods() {
        // TODO Implement registering dynamic methods to classes (optional)
    }

    void doWithApplicationContext() {
        // TODO Implement post initialization spring config (optional)
    }

    void onChange(Map<String, Object> event) {
        // TODO Implement code that is executed when any artefact that this plugin is
        // watching is modified and reloaded. The event contains: event.source,
        // event.application, event.manager, event.ctx, and event.plugin.
    }

    void onConfigChange(Map<String, Object> event) {
        // TODO Implement code that is executed when the project configuration changes.
        // The event is the same as for 'onChange'.
    }

    void onShutdown(Map<String, Object> event) {
        // TODO Implement code that is executed when the application shuts down (optional)
    }
}
