/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2016 PB <pbfordev@gmail.com>
// License:     MIT license
//////////////////////////////////////////////////////////////////////////// 


/**

@page page_introduction Introduction
 
 %wxAutoExcel is a C++ library which aims to make automating Microsoft Excel less painful. 
 It requires wxWidgets 3.1 or newer and works only on Microsft Windows. While it does not
 come even near to covering all the classes Microsoft Excel exposes, it still supports
 quite a bit of them. You can see what classes are supported <a href="annotated.html">here</a>.
 Be aware that even if a class is in the list, some of its methods or properties may not
 still be supported. %wxAutoExcel does not support any Microsot Excel events.
 
 It is assumed that the users of %wxAutoExcel are familar with VBA and Microsoft Excel model,
 as the library is only a bare C++ wrapper around it. But even if they are not, it could still 
 be possible to use %wxAutoExcel, as there are numerous resources about Microsoft Excel VBA
 programming on the internet. Moreover, very often it is possible to just record the macro
 inside Microsoft Excel and translate the resulting VBA code into corresponding %wxAutoExcel 
 calls - actually this is how I often write some of my code using %wxAutoExcel.

 */
