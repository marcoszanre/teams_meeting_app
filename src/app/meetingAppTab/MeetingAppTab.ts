import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/meetingAppTab/index.html")
@PreventIframe("/meetingAppTab/config.html")
@PreventIframe("/meetingAppTab/remove.html")
export class MeetingAppTab {
    
}
