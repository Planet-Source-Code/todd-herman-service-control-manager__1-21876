<div align="center">

## Service Control Manager


</div>

### Description

I have painstakingly researched everything dealing with services. It has taken me a while to weed through all the garbage that Microsoft dishes out but I finally did. This is a recreation (with a few tweaks) of the Service Control Manager for NT.

It doesn't do everything yet. This code shows you how to start/stop/pause/continue services. Beyond that, and even more important, it shows you how to determine IF a service can be stopped.. if it IS disabled.. if an error occurred, etc.

This is my first time submitting code to Planetsourcecode and I hope you guys vote for my code. I am planning on updating this app shortly to fully work with 2000 (using some new 2000 features) and enable the configuration of services. I also added a feature that Microsoft includes in there SCM. I am able to determine if a service depends on another service so all dependent services stop if you stop the main service. I can go into details and answer questions via feedback.

Please send me your feedback.
 
### More Info
 
This app deals with services so Windows NT is required. I also use Sheridan 3d controls (threed32.ocx) which comes on the VB cd (under tools) but is not a standard install.

If you don't know anything about services.. don't mess with them. Stopping certain services can cause problems with your system.

You also need to be an administrator of a machine to mess with services.


<span>             |<span>
---                |---
**Submitted On**   |2001-03-23 16:18:16
**By**             |[Todd Herman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/todd-herman.md)
**Level**          |Advanced
**User Rating**    |5.0 (85 globes from 17 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD174683232001\.zip](https://github.com/Planet-Source-Code/todd-herman-service-control-manager__1-21876/archive/master.zip)








