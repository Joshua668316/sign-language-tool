/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={34784:function(t,e,n){const r=n(99691);e.parseFont=r,e.createCanvas=function(t,e){return Object.assign(document.createElement("canvas"),{width:t,height:e})},e.createImageData=function(t,e,n){switch(arguments.length){case 0:return new ImageData;case 1:return new ImageData(t);case 2:return new ImageData(t,e);default:return new ImageData(t,e,n)}},e.loadImage=function(t,e){return new Promise((function(n,r){const o=Object.assign(document.createElement("img"),e);function i(){o.onload=null,o.onerror=null}o.onload=function(){i(),n(o)},o.onerror=function(){i(),r(new Error('Failed to load the image "'+t+'"'))},o.src=t}))}},99691:function(t){"use strict";const e="'([^']+)'|\"([^\"]+)\"|[\\w\\s-]+",n=new RegExp("(bold|bolder|lighter|[1-9]00) +","i"),r=new RegExp("(italic|oblique) +","i"),o=new RegExp("(small-caps) +","i"),i=new RegExp("(ultra-condensed|extra-condensed|condensed|semi-condensed|semi-expanded|expanded|extra-expanded|ultra-expanded) +","i"),a=new RegExp(`([\\d\\.]+)(px|pt|pc|in|cm|mm|%|em|ex|ch|rem|q) *((?:${e})( *, *(?:${e}))*)`),c={};t.exports=t=>{if(c[t])return c[t];const e=a.exec(t);if(!e)return;const u={weight:"normal",style:"normal",stretch:"normal",variant:"normal",size:parseFloat(e[1]),unit:e[2],family:e[3].replace(/["']/g,"").replace(/ *, */g,",")};let s,l,f,h;const p=t.substring(0,e.index);switch((s=n.exec(p))&&(u.weight=s[1]),(l=r.exec(p))&&(u.style=l[1]),(f=o.exec(p))&&(u.variant=f[1]),(h=i.exec(p))&&(u.stretch=h[1]),u.unit){case"pt":u.size/=.75;break;case"pc":u.size*=16;break;case"in":u.size*=96;break;case"cm":u.size*=96/2.54;break;case"mm":u.size*=96/25.4;break;case"%":break;case"em":case"rem":u.size*=16/.75;break;case"q":u.size*=96/25.4/4}return c[t]=u}},93180:function(t){"use strict";t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},5384:function(t,e,n){"use strict";t.exports=n.p+"046600ba6a6c802c622d.css"}},e={};function n(r){var o=e[r];if(void 0!==o)return o.exports;var i=e[r]={exports:{}};return t[r](i,i.exports,n),i.exports}n.m=t,n.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return n.d(e,{a:e}),e},n.d=function(t,e){for(var r in e)n.o(e,r)&&!n.o(t,r)&&Object.defineProperty(t,r,{enumerable:!0,get:e[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");if(r.length)for(var o=r.length-1;o>-1&&!t;)t=r[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,function(){"use strict";var t=n(34784).createCanvas;function e(t,e,n,r){this.imageSize=t,this.padding=e,this.textSpace=n,this.numPictures=r,this.width=t*r+e*(r+1),this.height=t+n}function r(t,e,n,r){var o=t.imageSize/Math.max(e,n),i=o*e,a=o*n,c=(t.imageSize-i)/2,u=(t.imageSize-a)/2;return{x:t.padding*r+t.imageSize*(r-1)+c,y:t.padding+u,imgWidth:i,imgHeight:a}}function o(t,e){return{x:t.padding*e+t.imageSize*(e-.5),y:.9*t.height}}function i(n,i){var a=new e(245,20,80,i.length),c=t(a.width,a.height),u=c.getContext("2d");return function(t,e,n,o){for(var i=1;i<=e.numPictures;i++)if(n.has(o[i-1])){var a=n.get(o[i-1]),c=r(e,a.naturalWidth,a.naturalHeight,i),u=c.x,s=c.y,l=c.imgWidth,f=c.imgHeight;t.drawImage(a,u,s,l,f)}}(u,a,n,i),function(t,e,n){t.fillStyle="#000000",t.font="48px Arial",t.textAlign="center",t.textBaseline="middle";for(var r=1;r<=e.numPictures;r++){var i=o(e,r),a=i.x,c=i.y;t.fillText(n[r-1],a,c)}}(u,a,i,n.size),c.toDataURL().split(",")[1]}function a(t){return a="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},a(t)}function c(){c=function(){return e};var t,e={},n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(t,e,n){t[e]=n.value},i="function"==typeof Symbol?Symbol:{},u=i.iterator||"@@iterator",s=i.asyncIterator||"@@asyncIterator",l=i.toStringTag||"@@toStringTag";function f(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(t){f=function(t,e,n){return t[e]=n}}function h(t,e,n,r){var i=e&&e.prototype instanceof w?e:w,a=Object.create(i.prototype),c=new T(r||[]);return o(a,"_invoke",{value:_(t,n,c)}),a}function p(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}e.wrap=h;var d="suspendedStart",m="suspendedYield",g="executing",y="completed",v={};function w(){}function x(){}function b(){}var E={};f(E,u,(function(){return this}));var L=Object.getPrototypeOf,S=L&&L(L(R([])));S&&S!==n&&r.call(S,u)&&(E=S);var O=b.prototype=w.prototype=Object.create(E);function I(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function P(t,e){function n(o,i,c,u){var s=p(t[o],t,i);if("throw"!==s.type){var l=s.arg,f=l.value;return f&&"object"==a(f)&&r.call(f,"__await")?e.resolve(f.__await).then((function(t){n("next",t,c,u)}),(function(t){n("throw",t,c,u)})):e.resolve(f).then((function(t){l.value=t,c(l)}),(function(t){return n("throw",t,c,u)}))}u(s.arg)}var i;o(this,"_invoke",{value:function(t,r){function o(){return new e((function(e,o){n(t,r,e,o)}))}return i=i?i.then(o,o):o()}})}function _(e,n,r){var o=d;return function(i,a){if(o===g)throw new Error("Generator is already running");if(o===y){if("throw"===i)throw a;return{value:t,done:!0}}for(r.method=i,r.arg=a;;){var c=r.delegate;if(c){var u=j(c,r);if(u){if(u===v)continue;return u}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===d)throw o=y,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=g;var s=p(e,n,r);if("normal"===s.type){if(o=r.done?y:m,s.arg===v)continue;return{value:s.arg,done:r.done}}"throw"===s.type&&(o=y,r.method="throw",r.arg=s.arg)}}}function j(e,n){var r=n.method,o=e.iterator[r];if(o===t)return n.delegate=null,"throw"===r&&e.iterator.return&&(n.method="return",n.arg=t,j(e,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),v;var i=p(o,e.iterator,n.arg);if("throw"===i.type)return n.method="throw",n.arg=i.arg,n.delegate=null,v;var a=i.arg;return a?a.done?(n[e.resultName]=a.value,n.next=e.nextLoc,"return"!==n.method&&(n.method="next",n.arg=t),n.delegate=null,v):a:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,v)}function k(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function z(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function T(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(k,this),this.reset(!0)}function R(e){if(e||""===e){var n=e[u];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,i=function n(){for(;++o<e.length;)if(r.call(e,o))return n.value=e[o],n.done=!1,n;return n.value=t,n.done=!0,n};return i.next=i}}throw new TypeError(a(e)+" is not iterable")}return x.prototype=b,o(O,"constructor",{value:b,configurable:!0}),o(b,"constructor",{value:x,configurable:!0}),x.displayName=f(b,l,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===x||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,b):(t.__proto__=b,f(t,l,"GeneratorFunction")),t.prototype=Object.create(O),t},e.awrap=function(t){return{__await:t}},I(P.prototype),f(P.prototype,s,(function(){return this})),e.AsyncIterator=P,e.async=function(t,n,r,o,i){void 0===i&&(i=Promise);var a=new P(h(t,n,r,o),i);return e.isGeneratorFunction(n)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},I(O),f(O,l,"Generator"),f(O,u,(function(){return this})),f(O,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},e.values=R,T.prototype={constructor:T,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=t,this.done=!1,this.delegate=null,this.method="next",this.arg=t,this.tryEntries.forEach(z),!e)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=t)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var n=this;function o(r,o){return c.type="throw",c.arg=e,n.next=r,o&&(n.method="next",n.arg=t),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var a=this.tryEntries[i],c=a.completion;if("root"===a.tryLoc)return o("end");if(a.tryLoc<=this.prev){var u=r.call(a,"catchLoc"),s=r.call(a,"finallyLoc");if(u&&s){if(this.prev<a.catchLoc)return o(a.catchLoc,!0);if(this.prev<a.finallyLoc)return o(a.finallyLoc)}else if(u){if(this.prev<a.catchLoc)return o(a.catchLoc,!0)}else{if(!s)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return o(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===t||"continue"===t)&&i.tryLoc<=e&&e<=i.finallyLoc&&(i=null);var a=i?i.completion:{};return a.type=t,a.arg=e,i?(this.method="next",this.next=i.finallyLoc,v):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),v},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),z(n),v}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;z(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(e,n,r){return this.delegate={iterator:R(e),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=t),v}},e}function u(t,e,n,r,o,i,a){try{var c=t[i](a),u=c.value}catch(t){return void n(t)}c.done?e(u):Promise.resolve(u).then(r,o)}function s(t){return function(){var e=this,n=arguments;return new Promise((function(r,o){var i=t.apply(e,n);function a(t){u(i,r,o,a,c,"next",t)}function c(t){u(i,r,o,a,c,"throw",t)}a(void 0)}))}}function l(){return f.apply(this,arguments)}function f(){return(f=s(c().mark((function t(){var e,n,r,o,i;return c().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return e=h(),n=document.getElementById("fileElem").files,r=Array.from(n).filter((function(t){return e.includes(t.name.split(".")[0])})),o=r.map((function(t){var e=new FileReader;return new Promise((function(n,r){e.onload=function(e){var o=new Image;o.onload=function(){return n({name:t.name.split(".")[0],image:o})},o.onerror=r,o.src=e.target.result},e.onerror=r,e.readAsDataURL(t)}))})),t.next=6,Promise.all(o);case 6:return i=t.sent,t.abrupt("return",new Map(i.map((function(t){return[t.name,t.image]}))));case 8:case"end":return t.stop()}}),t)})))).apply(this,arguments)}function h(){return document.getElementById("text-input").value.match(/(\b[^\s]+\b)/g)}function p(){return d.apply(this,arguments)}function d(){return(d=s(c().mark((function t(){var e,n;return c().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return e=h(),t.next=3,l();case 3:n=t.sent,r=i(n,e),Office.context.document.setSelectedDataAsync(r,{coercionType:Office.CoercionType.Image},(function(t){var e;t.status===Office.AsyncResultStatus.Failed&&(e="Error: "+t.error.message,document.getElementById("message").innerText=e)}));case 6:case"end":return t.stop()}var r}),t)})))).apply(this,arguments)}function m(t){return g.apply(this,arguments)}function g(){return(g=s(c().mark((function t(e){return c().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return document.getElementById("message").innerText="",t.next=3,e();case 3:case"end":return t.stop()}}),t)})))).apply(this,arguments)}Office.onReady((function(t){t.host===Office.HostType.PowerPoint&&(document.getElementById("app-body").style.display="flex",document.getElementById("insert-image").onclick=function(){return m(p)},document.getElementById("fileElem").onchange=function(){return m((function(t){return function(t){var e=document.getElementById("file_names"),n=document.getElementById("fileElem").files;e.textContent="";for(var r=0;r<n.length;r++)e.textContent+=n[r].name,r!==n.length-1&&(e.textContent+=", ")}()}))})}))}(),function(){"use strict";var t=n(93180),e=n.n(t),r=new URL(n(5384),n.b);e()(r)}()}();
//# sourceMappingURL=taskpane.js.map