let scatterLeft = 0, scatterTop = 0;
let scatterMargin = {top: 20, right: 30, bottom: 30, left: 100};
scatterWidth = 500 - scatterMargin.left - scatterMargin.right;
scatterHeight = 500 - scatterMargin.top - scatterMargin.bottom;

// append the svg object to the body of the page
var scatterSvg = d3.select("svg")
                    .append("g")
                    .attr("class", "scatter")
                    .attr("width", scatterWidth)
                    .attr("height", scatterHeight)
                    .attr("transform", "translate(" + scatterMargin.left + "," + scatterMargin.top + ")");
                    
var div = d3.select("body").append("div")
            .attr("class", "tooltip")
            .style("opacity", 0); 

function scatterChart(scatterData)
{
    // Add X axis
    var scatterX = d3.scaleLinear()					
                     .domain(d3.extent(scatterData, d => d.W)).nice()
                     .range([ 0, scatterWidth ]);
        scatterSvg.append("g")
                  .attr("transform", "translate(0," + scatterHeight + ")")
                  .call(d3.axisBottom(scatterX));
        scatterSvg.append("text")
                  .attr("transform", "translate(" + (scatterWidth/2 ) + " ," + (scatterHeight + scatterMargin.top + 20) + ")")
                  .style("text-anchor", "middle")
                  .text("W");

    // Add Y axis
    var scatterY = d3.scaleLinear()
                     .domain(d3.extent(scatterData, d => d.Gp_W)).nice()
                     .range([ scatterHeight, 0]);
        scatterSvg.append("g").call(d3.axisLeft(scatterY));
        scatterSvg.append("text")
                  .attr("transform", "rotate(-90)")
                  .attr("y", 5 - scatterMargin.left)
                  .attr("x",5 - (scatterHeight / 2))
                  .attr("dy", "1em")
                  .style("text-anchor", "middle")
                  .text("Gp_W");

    // Brush Event
    brush = scatterSvg.append("g")
                      .call(d3.brush().extent([[0, 0], [scatterWidth, scatterHeight]])
                      .on("brush", brushed)
                      .on("end", brushended)); // link bar, dist chart       
    
    // Add dots & mouse event
    scatterSvg.append('g')
              .selectAll("dot")
              .data(scatterData)
              .enter()
              .append("circle")
              .attr("r", 2)
              .style("fill", "#69b3a2")
              .attr("cx", function (d) { return scatterX(d.W); } )
              .attr("cy", function (d) { return scatterY(d.Gp_W); } )
              .attr("class", "non_brushed")
              .on("mouseover", function(d, i) {
                                                div.transition().duration(200).style("opacity", .9);
                                                div .html("W : " + d.W + "<br/>" + "Gp_W : " + d.Gp_W + "<br/>" + "Index : " + i)       
                                                .style("left", (d3.event.pageX) + "px")             
                                                .style("top", (d3.event.pageY - 28) + "px");
              })
              .on('mouseout', () => {
                                        div.transition().duration(100).style('opacity', 0);
              });

    function brushed() {
        
        var s = d3.event.selection,
            x0 = s[0][0],
            y0 = s[0][1],
            dx = s[1][0] - x0,
            dy = s[1][1] - y0;
        
        var circles = scatterSvg.selectAll("circle");
            circles.style("fill", function (d, i) {
                    if (scatterX(d.W) >= x0 && scatterX(d.W) <= x0 + dx && scatterY(d.Gp_W) >= y0 && scatterY(d.Gp_W) <= y0 + dy)
                    { 
                        d3.select(this).attr("class", "brushed").attr('id', i);
                        return "red"; 
                    }else {
                        d3.select(this).attr("class", "non_brushed"); 
                        return "#69b3a2"; 
                    }
            });
    }

    function brushended() {
        // Get brushed data
        var d_brushed =  d3.selectAll(".brushed").data(); 
        var brushed_select = document.querySelector(".brushed");

        console.log("W : " + d_brushed[0].W);
        console.log("Gp_W : " + d_brushed[0].Gp_W);
        console.log("index : " + brushed_select.getAttribute("id"));

        d3.select('.removeBtn')
          .on('click', function() {
              d3.selectAll("circle.brushed").remove();
        });        

        d3.select('.updateBtn')
          .on('click', function() {
              d3.set(scatterData).remove(brushed_select.getAttribute("id"));
        });
    }


}