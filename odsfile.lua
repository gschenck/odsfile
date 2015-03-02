-- module(...,package.seeall)
zip=require "minizip"
xmlparser = require ("luaxml-mod-xml")
handler = require("luaxml-mod-handler")

local function load(filename)
  local p = {
    file = zip.open(filename),
    content_file_name = "content.xml",
    loadContent = function(self,filename)
      local treehandler = handler.simpleTreeHandler()
      local filename = filename or self.content_file_name  
      local text
      if self.file:locate_file(filename) then
        text = self.file:extract(filename)
      else
        print(filename.." not found!")       
        exit()
      end
      local xml = xmlparser.xmlParser(treehandler)
      xml:parse(text)
      return treehandler
    end
  }
  return p
end

local function getTable(x,table_name)
  local tables = x.root["office:document-content"]["office:body"]["office:spreadsheet"]["table:table"]
  if #tables > 1 then
    if type(tables) == "table" and table_name ~= nil then 
        for k,v in pairs(tables) do
          if(v["_attr"]["table:name"]==table_name) then
            return v, k
          end 
        end
    elseif type(tables) == "table" and table_name == nil then
      return tables[1], 1  
    else 
      return tables  
    end
  else 
    return tables
  end
end

local function table_slice (values,i1,i2)
  -- Function from http://snippets.luacode.org/snippets/Table_Slice_116
  local res = {}
  local n = #values
  -- default values for range
  i1 = i1 or 1
  i2 = i2 or n
  if i2 < 0 then
    i2 = n + i2 + 1
  elseif i2 > n then
    i2 = n
  end
  if i1 < 1 or i1 > n then
    return {}
  end
  local k = 1
  for i = i1,i2 do
    res[k] = values[i]
    k = k + 1
  end
  return res
end

local function tableValues(tbl,x1,y1,x2,y2)
  local t= {}
  if type(tbl["table:table-row"])=="table" then
    local rows = table_slice(tbl["table:table-row"],y1,y2)
    for k,v in pairs(rows) do
      -- In every sheet, there are two rows with no data at the bottom, we need to strip them
      if(v["_attr"] and v["_attr"]["table:number-rows-repeated"] and tonumber(v["_attr"]["table:number-rows-repeated"])>10000) then break end
      local j = {}
      if #v["table:table-cell"] > 1 then
        local r = table_slice(v["table:table-cell"],x1,x2)
        for p,n in pairs(r) do
          local value=n["text:p"] 
          if type(value) == "table" then
            value = n["_attr"]["office:value"]
          elseif n["_attr"] and n["_attr"]["office:value-type"]=="float" and type(value)=="string" then
            value=value:gsub(",", ".")
        end
          if value and type(value) ~= "table" and value~="" then 
            table.insert(j,{["value"]=value, attr=n["_attr"]})
      else
            table.insert(j,{["value"]="", attr=""})
          end
        end
      else
        local value=v["table:table-cell"]["text:p"]
        if type(value) == "table" then
            value = v["_attr"]["office:value"]
        elseif v["_attr"]["office:value-type"]=="float" and type(value)=="string" then
          value:gsub(",", ".")
        end  
        if value and type(value) ~= "table" and  value~="" then 
          local p = {["value"]=value,attr=v["table:table-cell"]["_attr"]} 
        table.insert(j,p) 
        else
          table.insert(j,{["value"]="", attr=""})
        end
      end  
      if #j>0 then table.insert(t,j) end
    end
  end
  return t
end

local function getRange(range)
  local r = range:lower()
  local function getNumber(s)
    if s == "" or s == nil then return nil end
    local f,ex = 0,0
    for i in string.gmatch(s:reverse(),"(.)") do
      f = f + (i:byte()-96) * 26 ^ ex
      ex = ex + 1 
    end
    return f
  end
  for x1,y1,x2,y2 in r:gmatch("(%a*)(%d*):*(%a*)(%d*)") do
    return getNumber(x1),tonumber(y1),getNumber(x2),tonumber(y2) 
   --print(string.format("%s, %s, %s, %s",getNumber(x1),y1,getNumber(x2),y2))
  end
end

local function interp(s, tab)
  return (s:gsub('(-%b{})', 
    function(w) 
      s = w:sub(3, -2)
      s = tonumber(s) or s
      return tab[s] or w 
    end)
  )
end


-- Interface for adding new rows to the spreadsheet
local function newRow()
  local p = {
    pos = 0,
    cells = {},
    -- Generic function for inserting cell
    addCell = function(self,val, attr,pos)
      if pos then
        table.insert(self.cells,pos,{["text:p"] = val, ["_attr"] = attr})
        self.pos = pos
      else
        self.pos = self.pos + 1
        table.insert(self.cells,self.pos,{["text:p"] = val, ["_attr"] = attr})
      end
    end, 
    addString = function(self,s,attr,pos)
      local attr = attr or {}
      attr["office:value-type"] = "string"
      self:addCell(s,attr,pos)
    end,
    addFloat = function(self,i,attr,pos)
      local attr = attr or {}
      local s = tonumber(i) or 0
      s = tostring(s)
      attr["office:value-type"] = "float"
      attr["office:value"] = s
      self:addCell(s,attr,pos)
    end, 
    findLastRow = function(self,sheet)
      for i= #sheet["table:table-row"],1,-1 do
        if sheet["table:table-row"][i]["_attr"]["table:number-rows-repeated"] then
          return i
        end
      end
      return #sheet["table:table-row"]+1
    end,
    insert = function(self, sheet, pos)
      local t = {}
      local pos = pos or self:findLastRow(sheet)
      print("pos je: ",pos)
      if sheet["table:table-column"]["_attr"] and sheet["table:table-column"]["_attr"]["table:number-columns-repeated"] then
	table_columns = sheet["table:table-column"]["_attr"]["table:number-columns-repeated"]
      else 
	table_columns = #sheet["table:table-column"]
      end
      for i=1, table_columns do
        table.insert(t,self.cells[i] or {})  
      end
      t = {["table:table-cell"]=t}
      table.insert(sheet["table:table-row"],pos,t)
    end
  }
  return p
end


-- function for updateing the archive. Depends on external zip utility
local function updateZip(zipfile, updatefile)
  local command  =  string.format("zip %s %s",zipfile, updatefile)
  print ("Updating an ods file.\n" ..command .."\n Return code: ", os.execute(command))  
end

-- set up a module
local M={}
M.load=load
M.getTable = getTable
M.table_slice = table_slice
M.tableValues = tableValues
M.getRange = getRange
M.interp = interp
M.newRow = newRow
M.updateZip = updateZip
return M