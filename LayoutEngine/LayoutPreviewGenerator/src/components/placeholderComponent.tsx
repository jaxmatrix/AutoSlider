import { useEffect, useLayoutEffect, useState } from "react";
import { useSearchParams } from "react-router-dom";

interface ContentJson { 
    [key: string]: { 
        content: string; 
        type: string; 
    } 
};

interface LayoutJson { 
    layout: { 
        cols: number; 
        rows: number; 
    }; 
    children: { 
        [key: string]: { 
            row: number; 
            col: number; 
            rowspan: number; 
            colspan: number; 
        }; 
    }; 
};

const testLayout =   {
    layout: { cols: 12, rows: 8 },
    children: {
      "1": { row: 1, col: 1, rowspan: 1, colspan: 12 },
      "2": { row: 2, col: 1, rowspan: 1, colspan: 8 },
      "3": { row: 2, col: 9, rowspan: 1, colspan: 4 },
      "4": { row: 3, col: 1, rowspan: 1, colspan: 12 },
      "5": { row: 4, col: 1, rowspan: 1, colspan: 4 },
      "6": { row: 4, col: 5, rowspan: 1, colspan: 8 },
      "7": { row: 5, col: 1, rowspan: 1, colspan: 12 },
      "8": { row: 6, col: 1, rowspan: 1, colspan: 6 },
      "9": { row: 6, col: 7, rowspan: 1, colspan: 6 },
      "10":{ row: 7, col: 1, rowspan: 1, colspan: 4 },
      "11":{ row: 7, col: 5, rowspan: 1, colspan: 8 }
    }
}  

const testContent ={
    "1":  { content: "Product", type: "header" },
    "2":  { content: "First of its kind", type: "sub-header" },
    "3":  { content: "Novel Technology", type: "sub-header" },
    "4":  { content: "Instant Result", type: "highlight" },
    "5":  { content: "Simplified Compliances", type: "text" },
    "6":  { content: "Designed for frequent use", type: "text" },
    "7":  { content: "Single Biomarker Testing at INR 70", type: "highlight" },
    "8":  { content: "Ease of Use", type: "text" },
    "9":  { content: "Drop 200ul to 500ul of saliva on chip", type: "text" },
    "10": { content: "Ultrasensitive Detection", type: "highlight" },
    "11": { content: "Get Instant results and medical advice on your phone", type: "text" }
  } 


const PlaceholderComponent = () => {
    const [searchParams] = useSearchParams();
    const query = searchParams.get('layout');
    const [data, setData] = useState("Data Is not Set");
    const [layout, setLayout] = useState<LayoutJson>({
        layout : { rows : 1, cols : 1},
        children : {}
    })
    const [content, setContent] = useState<ContentJson>({})
    const [height, setHeight] = useState(123)
    const [width, setWidth] = useState(600)

    useEffect(()=>{
        const handleMessage = (e) => {
            if (e.origin === 'http://localhost:4000') {
                setData(e.data);
            }
        };

        window.addEventListener('message', handleMessage);
        return () => window.removeEventListener('message', handleMessage)
    }, [])

    useEffect(() => {
        setLayout(testLayout) 
        console.log("Setting Test Layout", layout, testLayout.layout)

        setContent(testContent)
        console.log("Setting Test Content", content, testContent)


    }, [])

    useLayoutEffect(()=>{
        setHeight(window.innerHeight)
        setWidth(window.innerWidth)
    }, [])

    if (true) {
        return (
            <div style={{
                height: 1080,
                width : 1920,
                padding: 50, 
                transform:`scale(${width*0.8/1920})` ,
                transformOrigin: 'top left',
                position : "relative"
            }} className={` gap-2 grid grid-cols-${layout.layout.cols} grid-rows-${layout.layout.rows}`}>
                { Object.keys(layout.children).map((contentId)=>(
                    <div key={contentId} 
                        className={
                            `border rounded-2xl p-4 bg-red-100` + 
                            ` col-start-${layout.children[contentId].col}` +
                            ` row-start-${layout.children[contentId].row}` +
                            ` col-span-${layout.children[contentId].colspan}` +
                            ` row-span-${layout.children[contentId].rowspan}` 
                    }>
                        { content[contentId].content }
                    </div>
                )) 
                } 

            </div>
        );
    } else {
        return (
            <div className="grid grid-cols-12 height-[600px] grid-rows-8 gap-2">
                <div className="bg-red-100 border col-start-1 col-span-12 row-start-1 row-span-1">1</div>
                <div className="bg-red-100 border col-start-1 col-span-8  row-start-2 row-span-1">2</div>
                <div className="bg-red-100 border col-start-9 col-span-4 row-start-2 row-span-1">3</div>
                <div className="bg-red-100 border col-start-1 col-span-12 row-start-3 row-span-1">4</div>
                <div className="bg-red-100 border col-start-1 col-span-4 row-start-4 row-span-1">5</div>
                <div className="bg-red-100 border col-start-5 col-span-8 row-start-4 row-span-1">6</div>
                <div className="bg-red-100 border col-start-1 col-span-12 row-start-5 row-span-1">7</div>
                <div className="bg-red-100 border col-start-1 col-span-6 row-start-6 row-span-1">8</div>
                <div className="bg-red-100 border col-start-7 col-span-6 row-start-6 row-span-1">9</div>
                <div className="bg-red-100 border col-start-1 col-span-4 row-start-7 row-span-1">10</div>
                <div className="bg-red-100 border col-start-5 col-span-8 row-start-7 row-span-1">11</div>
            </div>
        )
    }
}

export { PlaceholderComponent }