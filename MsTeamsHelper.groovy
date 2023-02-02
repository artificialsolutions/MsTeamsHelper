/*
   Copyright 2022 Artificial Solutions
   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0
       
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

/*
Name: MsTeamsHelper
Description: Helper class for Teams connector 
*/

public class MsTeamsHelper {

    // Adaptive card headers
	private static adaptiveCard = [:]
	public MsTeamsHelper(version=""){
	    adaptiveCard = new HashMap<>()
	    adaptiveCard.type = "AdaptiveCard"
	    adaptiveCard.version = version?:"1.4"
		adaptiveCard.body = []
		adaptiveCard.actions = []
	}
	
	// Create Output Parameters
	public static void createOutputParamters(def _){
		
		// Search for simple OPs and convert to JSON
		List<String> order = []
        Iterator it = _.getOutputParameters().entrySet().iterator()
        while(it.hasNext()){
			def entry = it.next()
			
			switch(entry.getKey()){
				case "teams_order":
					order = entry.getValue().trim().split(",")
					it.remove()
					break;
				case "teams_buttons":
					addButtons(entry.getValue())
					it.remove()
					break;
				case "teams_image":
					addImage(entry.getValue())
					it.remove()
					break;
				case "teams_video":
					addMediaFile(entry.getValue(),"video")
					it.remove()
					break;
				case "teams_audio":
					addMediaFile(entry.getValue(),"audio")
					it.remove()
					break;
				case "teams_webvideo":
					addWebVideo(entry.getValue())
					it.remove()
					break;
				case "teams_links":
					addLinks(entry.getValue())
					it.remove()
					break;
				case "teams_text":
					addTextBlock(entry.getValue())
					it.remove()
					break;
			}
		}
		
		if (order) adjustOrder(order)
		if (adaptiveCard.body||adaptiveCard.actions) _.putOutputParameter("msbotframework",createJson())
		
	}
	
	// Adjust the order of body and actions
	private static void adjustOrder(List<String> order){
		
		def tempBody = []
		def tempActions = []
		for (messageType in order){

			if (messageType == "image"){
				Iterator it = adaptiveCard.body.iterator()
				while(it.hasNext()){
					def content = it.next()
					if (content.get("type") == "Image"){
						tempBody << content
						it.remove()
					}
				}
			}
			else if (messageType == "text"){
				Iterator it = adaptiveCard.body.iterator()
				while(it.hasNext()){
					def content = it.next()
					if (content.get("type") == "TextBlock"){
						tempBody << content
						it.remove()
					}
				}
			}
			else if (messageType == "media"){
				Iterator it = adaptiveCard.body.iterator()
				while(it.hasNext()){
					def content = it.next()
					if (content.get("type") == "Media"){
						tempBody << content
						it.remove()
					}
				}
			}
			else if (messageType == "buttons"){
				Iterator it = adaptiveCard.actions.iterator()
				while(it.hasNext()){
					def content = it.next()
					if (content.get("type") == "Action.Submit"){
						tempActions << content
						it.remove()
					}
				}
			}
			else if (messageType == "links"){
				Iterator it = adaptiveCard.actions.iterator()
				while(it.hasNext()){
					def content = it.next()
					if (content.get("type") == "Action.OpenUrl"){
						tempActions << content
						it.remove()
					}
				}
			}
		}
		for(content in tempBody) adaptiveCard.body << content
		for(content in tempActions) adaptiveCard.actions << content
		
	}	
	
	// Create JSON
	private static String createJson(){
		return new groovy.json.JsonBuilder(adaptiveCard).toString()
	}

	// Create Text Block
	private static void addTextBlock(String textOP){
		
		try {
			def text = []
			def size = []
			if(textOP.contains("|")){
				text = textOP.split("\\|")[0]
				size = textOP.split("\\|")[1]
			} else {
				text = textOP
			}
			def textBlock = [:]
			textBlock.type = "TextBlock"
			textBlock.text = text
			textBlock.size = size?:"medium"
			adaptiveCard.body << textBlock 
        } catch (Exception e){
            println "Error addTextBlock: " + e.getMessage()
        }
		
	}

	// Buttons and links
	private static Map createButton(String item){
		
		def action = [:]
		action.type = "Action.Submit"
		action.title = item
		action.data = [:]
		action.data.buttonChoice = item
		action.data.msteams = [:]
		action.data.msteams.type = "messageBack"
		action.data.msteams.displayText = item
		return action
		
	}
		
    private static void addButtons(String actionsOP){
      
		try {
			def actionList = actionsOP.split("\\|")
			if (actionList.size()<=6){
				for (item in actionList) 
					adaptiveCard.actions << createButton(item)
			}
			else if (actionList.size()<12){
				
				// Split the button list 
				def sublist1 = Arrays.copyOfRange(actionList,0,5)
				def sublist2 = Arrays.copyOfRange(actionList,5,actionList.size())
				
				// Add the buttons in list 1
				for (item in sublist1){
					adaptiveCard.actions << createButton(item)
				}
				
				// Create embedded card to carry buttons in list 2
				def showCard = [:]
				showCard.type = "Action.ShowCard"
				showCard.title = "More"
				showCard.card = [:]
				showCard.card.type = "AdaptiveCard"
				showCard.card.actions = []
				for (item in sublist2){
					showCard.card.actions << createButton(item)
				}
				adaptiveCard.actions << showCard	
				
			}
			else {
				throw new Exception("Too many buttons. Please create less than 12 buttons at the same time.")
			}
		} catch (Exception e){
            println "Error addButtons: " + e.getMessage()
        }

    }
	
	private static void addLinks(String linksOP){
      
		try {
			def links = linksOP.split("\\|")
			for (link in links){
				def action = [:]
				action.type = "Action.OpenUrl"
				action.title = link.split(",")[0]
				action.url = link.split(",")[1]
				adaptiveCard.actions << action
			}
		} catch (Exception e){
            println "Error addLinks: " + e.getMessage()
        }

    }

    // Media attachments
	
	// Create mime types list
	private static mimeTypes = [:]
	static {
		
		mimeTypes.audio = [:]
		mimeTypes.video = [:]
	
		// Audios
		mimeTypes.audio.put("3gp","audio/3gpp")
		mimeTypes.audio.put("3g2","audio/3gpp2")
		mimeTypes.audio.put("adp","audio/adpcm")
		mimeTypes.audio.put("aiff","audio/aiff")
		mimeTypes.audio.put("aif","audio/aiff")
		mimeTypes.audio.put("aff","audio/aiff")
		mimeTypes.audio.put("au","audio/basic")
		mimeTypes.audio.put("snd","audio/basic")
		mimeTypes.audio.put("flac","audio/flac")
		mimeTypes.audio.put("kar","audio/midi")
		mimeTypes.audio.put("mid","audio/midi")
		mimeTypes.audio.put("midi","audio/midi")
		mimeTypes.audio.put("rmi","audio/midi")
		mimeTypes.audio.put("mp4a","audio/mp4")
		mimeTypes.audio.put("m2a","audio/mpeg")
		mimeTypes.audio.put("m3a","audio/mpeg")
		mimeTypes.audio.put("mp2","audio/mpeg")
		mimeTypes.audio.put("mp2a","audio/mpeg")
		mimeTypes.audio.put("mp3","audio/mpeg")
		mimeTypes.audio.put("mpga","audio/mpeg")
		mimeTypes.audio.put("oga","audio/ogg")
		mimeTypes.audio.put("ogg","audio/ogg")
		mimeTypes.audio.put("spx","audio/ogg")
		mimeTypes.audio.put("opus","audio/opus")
		mimeTypes.audio.put("eol","audio/vnd.digital-winds")
		mimeTypes.audio.put("dts","audio/vnd.dts")
		mimeTypes.audio.put("dtshd","audio/vnd.dts.hd")
		mimeTypes.audio.put("lvp","audio/vnd.lucent.voice")
		mimeTypes.audio.put("pya","audio/vnd.ms-playready.media.pya")
		mimeTypes.audio.put("wav","audio/wav")
		mimeTypes.audio.put("weba","audio/webm")
		mimeTypes.audio.put("aac","audio/x-aac")
		mimeTypes.audio.put("mka","audio/x-matroska")
		mimeTypes.audio.put("m3u","audio/x-mpegurl")
		mimeTypes.audio.put("wax","audio/x-ms-wax")
		mimeTypes.audio.put("wma","audio/x-ms-wma")
		mimeTypes.audio.put("ra","audio/x-pn-realaudio")
		mimeTypes.audio.put("ram","audio/x-pn-realaudio")
		mimeTypes.audio.put("rmp","audio/x-pn-realaudio-plugin")
	
		// Videos 
		mimeTypes.video.put("3gp","video/3gpp")
		mimeTypes.video.put("3g2","video/3gpp2")
		mimeTypes.video.put("h261","video/h261")
		mimeTypes.video.put("h263","video/h263")
		mimeTypes.video.put("h264","video/h264")
		mimeTypes.video.put("jpgv","video/jpeg")
		mimeTypes.video.put("jpgm","video/jpm")
		mimeTypes.video.put("jpm","video/jpm")
		mimeTypes.video.put("mj2","video/mj2")
		mimeTypes.video.put("mjp2","video/mj2")
		mimeTypes.video.put("mp4","video/mp4")
		mimeTypes.video.put("mp4v","video/mp4")
		mimeTypes.video.put("mpg4","video/mp4")
		mimeTypes.video.put("m1v","video/mpeg")
		mimeTypes.video.put("m2v","video/mpeg")
		mimeTypes.video.put("mpa","video/mpeg")
		mimeTypes.video.put("mpe","video/mpeg")
		mimeTypes.video.put("mpeg","video/mpeg")
		mimeTypes.video.put("mpg","video/mpeg")
		mimeTypes.video.put("ogv","video/ogg")
		mimeTypes.video.put("mov","video/quicktime")
		mimeTypes.video.put("qt","video/quicktime")
		mimeTypes.video.put("fvt","video/vnd.fvt")
		mimeTypes.video.put("m4u","video/vnd.mpegurl")
		mimeTypes.video.put("mxu","video/vnd.mpegurl")
		mimeTypes.video.put("pyv","video/vnd.ms-playready.media.pyv")
		mimeTypes.video.put("viv","video/vnd.vivo")
		mimeTypes.video.put("webm","video/webm")
		mimeTypes.video.put("f4v","video/x-f4v")
		mimeTypes.video.put("fli","video/x-fli")
		mimeTypes.video.put("flv","video/x-flv")
		mimeTypes.video.put("m4v","video/x-m4v")
		mimeTypes.video.put("mkv","video/x-matroska")
		mimeTypes.video.put("asf","video/x-ms-asf")
		mimeTypes.video.put("asx","video/x-ms-asf")
		mimeTypes.video.put("wm","video/x-ms-wm")
		mimeTypes.video.put("wmv","video/x-ms-wmv")
		mimeTypes.video.put("wmx","video/x-ms-wmx")
		mimeTypes.video.put("wvx","video/x-ms-wvx")
		mimeTypes.video.put("avi","video/x-msvideo")
		mimeTypes.video.put("movie","video/x-sgi-movie")
	
	}
	
	// Get Media file extension
	private static String getExtension(def filename = ""){
		return filename.replaceAll(".*\\.","")
	}
	
	// Get Mime Type
	private static String getMimeType(def url = "", def category = ""){
		return mimeTypes.get(category).get(getExtension(url))
	}
	
	// Add image
    private static void addImage(String imageOP) {
        
        try {
            def image = [:]
			def url = ""
			def alt = ""
			if(imageOP.contains("|")){
				url = imageOP.split("\\|")[0]
				alt = imageOP.split("\\|")[1]
			} else {
				url = imageOP
			}
			image.type = "Image"
			image.url = url
			image.altText = alt?:"This is an image"
			adaptiveCard.body << image
        } catch (Exception e){
            println "Error addImage: " + e.getMessage()
        }

    }
	
	// Add media file (mime type required)
    private static void addMediaFile(String mediaOP, String category) {
        
        try {
            def media = [:]
			def mediaSource = [:]
			def url = ""
			def alt = ""
			if(mediaOP.contains("|")){
				url = mediaOP.split("\\|")[0]
				alt = mediaOP.split("\\|")[1]
			} else {
				url = mediaOP
			}
			media.type = "Media"
			media.sources = []
			media.altText = alt?:"This is a media file"
			mediaSource.mimeType = getMimeType(url, category)
			mediaSource.url = url
			media.sources << mediaSource
			adaptiveCard.body << media
        } catch (Exception e){
            println "Error addMediaFile: " + e.getMessage()
        }

    }
	
	// Add web video (mime type omitted)	
    private static void addWebVideo(String mediaOP) {
        
        try {
            def media = [:]
			def mediaSource = [:]
			def url = ""
			def alt = ""
			if(mediaOP.contains("|")){
				url = mediaOP.split("\\|")[0]
				alt = mediaOP.split("\\|")[1]
			} else {
				url = mediaOP
			}
			media.type = "Media"
			media.sources = []
			media.altText = alt?:"This is a video"
			mediaSource.url = url
			media.sources << mediaSource
			adaptiveCard.body << media
        } catch (Exception e){
            println "Error addWebVideo: " + e.getMessage()
        }

    }
	
}