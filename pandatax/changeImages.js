console.log(Math.random());


(function(){
	let imgs = document.getElementsByTagName('source');
	for (let i = 0; i < imgs.length; i++) {
		imgs[i].srcset = "https://i.guim.co.uk/img/media/90d3278ac3c565b91d77fd8ffc4b0150827343f8/0_162_5578_3347/master/5578.jpg?width=300&quality=85&auto=format&fit=max&s=9c6eb847ca8bfb129b2fde3a93149277";
	}	
})();

(function(){
	let imgs = document.getElementsByTagName('img');
	for (let i = 0; i < imgs.length; i++) {
		imgs[i].src = "//i.imgur.com/fSgnUKWb.jpg"
		imgs[i].srcset = "//i.imgur.com/fSgnUKWb.jpg"
	}
})();


