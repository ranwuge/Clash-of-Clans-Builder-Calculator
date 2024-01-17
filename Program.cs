﻿using OfficeOpenXml;
using System.Collections;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Xml.Serialization;

internal class Program
{
    private static void Main(string[] args)
    {
        /*
        Time_Chunk_Pool tcp = new Time_Chunk_Pool();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string adress = @"D:\COC-Calculator\BuilderCalculator\defense_data.xlsx";
        ExcelPackage basic_excel = new ExcelPackage(adress);
        ExcelWorksheet defense_data = basic_excel.Workbook.Worksheets[1];
        ExcelWorksheet input_data = basic_excel.Workbook.Worksheets[0];
        tcp.From_input_to_chunk(input_data, defense_data);  //从用户输入构建时块池
        Console.WriteLine($"Start Programming...________________________________");
        foreach (Time_Chunk chunk in tcp.time_chunk_pool) 
        {
            Console.WriteLine($"name:{chunk.chunk_name},id:{chunk.chunk_id}," +
                                      $"cost_time:{chunk.chunk_time_length},level:{chunk.chunk_level}");  
        }
        Console.WriteLine($"length of the pool:{tcp.time_chunk_pool.Count}");
        */
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string adress = @"D:\COC-Calculator\BuilderCalculator\defense_data.xlsx";
        ExcelPackage basic_excel = new ExcelPackage(adress);
        ExcelWorksheet defense_data = basic_excel.Workbook.Worksheets[1];
        ExcelWorksheet input_data = basic_excel.Workbook.Worksheets[0];

        //根据输入生成建筑单元池
        Building_Unit_Pool bup =  Generator.Building_unit_pool_generator(input_data, defense_data, 12);
        Console.WriteLine($"In bup there is {bup.bu_set[0]}");
        //对建筑单元池中每一建筑单元生成其对应的时块集
        foreach (Building_Unit bu in bup.bu_set) 
        {
            Generator.Time_chunk_generator(bu, defense_data);
        }
        Builder_Pipeline_Pool bpp = new Builder_Pipeline_Pool(6);
        bup.total_time = bup.Get_Total_Time();

        while (bup.total_time != bpp.total_time) 
        {
            Builder_Pipeline bp = Generator.Find_Bp(bpp);
            Time_Chunk tc = Generator.Chunk_Allocator(bup, bpp);
            bpp.total_time += tc.chunk_time_length;
            Console.WriteLine($"BPP total time is {bpp.total_time}.");
            if (tc.chunk_name == "nothing") { break; }
        }
        Console.WriteLine($"______________________________________________________\n Statistic\n______________________________________________________");
        foreach (Builder_Pipeline bp in bpp.bp_set) 
        {
            Console.WriteLine($"Builder{bp.pipeline_id}'s total time is {bp.time_length}.");
        }
        Console.WriteLine($"______________________________________________________");
        Console.WriteLine($"BUP total time is {bup.Get_Total_Time()}.");
    }
}

class Time_Chunk :IDisposable
{
    public decimal chunk_time_length = 0;
    public string? chunk_name;
    public int chunk_level;
    public int chunk_id;
 
    public void Dispose()
    {
        throw new NotImplementedException();
    }
}
/*
class Time_Chunk_Pool
{
    public ArrayList time_chunk_pool = new();
    public int TH_num = 12;
    public void Add_Chunk_to_pool(Time_Chunk time_chunk)
    {
        time_chunk_pool.Add(time_chunk);
    }
    public void From_input_to_chunk(ExcelWorksheet input_data, ExcelWorksheet basic_data)
    {
        Time_Chunk added_chunk = new Time_Chunk();
        var input_row_Num = input_data.Dimension.Rows;
        var basic_row_Num = basic_data.Dimension.Rows;
        int max_level = 0;
        int building_id = 1;

        foreach (var cell in input_data.Cells[$"A2:A{input_row_Num}"]) 
        {
            //用户输入的建筑数据：
            string building_name = (string)cell.Value;
            Double building_level_D = (Double)cell.Offset(0, 1).Value;
            int building_level = (int)building_level_D ;

            //根据时块池，确定该建筑的id           
            if (cell.Start.Row == 3) { building_id = 1; }
            else
            {
                string a = cell.Offset(-1, 0).GetValue<string>();
                if (cell.Offset(-1, 0).GetValue<string>() != building_name) { building_id = 1; } 
                else { building_id++; }
            }
            //根据大本营等级确定该建筑可以升到的最大等级
            foreach (var b_cell in basic_data.Cells[$"A2:A{basic_row_Num}"])
            {
                string b_cell_name = (string)b_cell.Value;
                Double b_cell_levelD = (Double)b_cell.Offset(0, 1).Value;
                Double b_cell_THD = (Double)b_cell.Offset(0, 3).Value;
                Double? b_cell_THD_next = 100;
                if (b_cell.Start.Row != basic_row_Num) { b_cell_THD_next = (Double)b_cell.Offset(1, 3).Value; }
                else { b_cell_THD_next = (Double)b_cell.Offset(0, 3).Value; }
                int b_cell_level = (int)b_cell_levelD;
                int b_cell_TH = (int)b_cell_THD;
                int b_cell_TH_next = (int)b_cell_THD_next;

                if (building_name == b_cell_name && TH_num == b_cell_TH && building_level != b_cell_TH_next)
                {
                    max_level = b_cell_level;
                    break;
                }
            }

            //根据可以升级的等级添加时块
            for (int i = building_level + 1; i <= max_level; i++) 
            {
                added_chunk.chunk_level = i;
                //查找该等级下的消耗时间
                foreach (var b_cell in basic_data.Cells[$"A2:A{basic_row_Num}"])
                {
                    string b_cell_name = (string)b_cell.Value;
                    Double b_cell_level = (Double)b_cell.Offset(0, 1).Value;
                    Double b_cell_cost_time = (Double)b_cell.Offset(0, 2).Value;
                    if (building_name == b_cell_name && i == b_cell_level)
                    {
                        added_chunk.chunk_time_length = (decimal)b_cell_cost_time;
                        added_chunk.chunk_id = building_id;
                        added_chunk.chunk_name = building_name;
                        break;
                    }   
                }
                Time_Chunk ck = new Time_Chunk(added_chunk.chunk_time_length, added_chunk.chunk_name,
                                                added_chunk.chunk_level,added_chunk.chunk_id);
                Add_Chunk_to_pool(ck);
                Console.WriteLine($"name:{added_chunk.chunk_name},id:{added_chunk.chunk_id}," +
                                  $"cost_time:{added_chunk.chunk_time_length},level:{added_chunk.chunk_level}");
            }
        }
    }
}
*/
public class Building_Unit 
{
    public int bu_level;
    public int bu_id;
    public int bu_max_level;
    public decimal avalible_time;
    public string? bu_name;
    public ArrayList tc_set = new ArrayList();

}

class Building_Unit_Pool
{
    public decimal total_time = 0;
    public ArrayList bu_set = new ArrayList();
    public void Add_bu(Building_Unit bu) { bu_set.Add(bu); }
    public decimal Get_Total_Time() 
    {
        if (bu_set != null)
        {
            foreach (Building_Unit bu in bu_set)
            {
                foreach (Time_Chunk tc in bu.tc_set)
                {
                    total_time += tc.chunk_time_length;
                }
            }
            return total_time;
        }
        return total_time;
    }
}

static class Generator
{
    //这个方法用于从用户输入生成建筑单元池
    public static Building_Unit_Pool Building_unit_pool_generator(ExcelWorksheet input_data, ExcelWorksheet basic_data, int TH_level)
    {
        //构造建筑单元池
        Building_Unit_Pool bup = new Building_Unit_Pool();
        var input_row_Num = input_data.Dimension.Rows;
        var basic_row_Num = basic_data.Dimension.Rows;

        //初始化建筑单元数据
        int bu_max_level = 1;
        int bu_level;
        int bu_id = 1;
        string? bu_name;

        //根据用户输入的数据，补充建筑单元的数据：
        foreach (var cell in input_data.Cells[$"A2:A{input_row_Num}"])
        {
            bu_name = cell.GetCellValue<string>();  //建筑的名字
            bu_level = cell.Offset(0, 1).GetCellValue<int>();   //建筑的当前等级
            //建筑的id（与上一个建筑比较找出id）
            if (cell.Start.Row == 3)    
            { 
                bu_id = 1; 
            }
            else
            {
                string a = cell.Offset(-1, 0).GetValue<string>();
                if (cell.Offset(-1, 0).GetValue<string>() != bu_name) { bu_id = 1; }
                else { bu_id++; }
            }
            //建筑可升级到的最大等级(与下一个建筑的等级比较找出最大等级)
            foreach (var b_cell in basic_data.Cells[$"A2:A{basic_row_Num}"])
            {
                string b_cell_name = (string)b_cell.Value;
                int b_cell_level = b_cell.Offset(0, 1).GetCellValue<int>();
                int b_cell_TH = b_cell.Offset(0, 3).GetCellValue<int>();
                int b_cell_TH_next = 100;
                if (b_cell.Start.Row != basic_row_Num) { b_cell_TH_next = b_cell.Offset(1, 3).GetCellValue<int>(); }
                else { b_cell_TH_next = b_cell.Offset(0, 3).GetCellValue<int>(); }
                if (bu_name == b_cell_name && TH_level == b_cell_TH && TH_level != b_cell_TH_next)
                {
                    bu_max_level = b_cell_level;
                    break;
                }
            }
            //将建筑单元添加至建筑单元池中
            Building_Unit added_bu = new Building_Unit();        
            added_bu.bu_id = bu_id;
            added_bu.bu_name = bu_name;
            added_bu.bu_level = bu_level;
            added_bu.bu_max_level = bu_max_level;
            bup.Add_bu(added_bu);
            Console.WriteLine($"ADDing Bu: name:{added_bu.bu_name},id:{added_bu.bu_id}," +
                                    $"current_level:{added_bu.bu_level},max_level:{added_bu.bu_max_level}");           
        }
            return bup;
    }

    //这个方法用于补充建筑单元缺失的部分
    public static void Time_chunk_generator(Building_Unit bu, ExcelWorksheet basic_data)
    {

        var basic_row_Num = basic_data.Dimension.Rows;
        string? bu_name = bu.bu_name;
        int bu_level = bu.bu_level;
        int bu_id = bu.bu_id;
        int max_level = bu.bu_max_level;


        //查找所有时块
        for (int i = bu_level + 1; i <= max_level; i++)
        {
            Time_Chunk tc = new Time_Chunk();   //构建要用的时块
            tc.chunk_name = bu_name;    //该时块的名字
            tc.chunk_level = i;    //该时块的名字
            //查找所有时块的消耗时间
            foreach (var cell in basic_data.Cells[$"A2:A{basic_row_Num}"])
            {
                string b_cell_name = (string)cell.Value;
                Double b_cell_level = (Double)cell.Offset(0, 1).Value;
                Double b_cell_cost_time = (Double)cell.Offset(0, 2).Value;
                if (bu_name == b_cell_name && i == b_cell_level)
                {
                    tc.chunk_time_length = (decimal)b_cell_cost_time;
                    tc.chunk_time_length = Math.Round(tc.chunk_time_length,2);
                    tc.chunk_id = bu_id;
                    break;
                }
            }
            bu.tc_set.Add(tc);
            Console.WriteLine($"ADDing tc: name:{tc.chunk_name},id:{tc.chunk_id}," +
                      $"current_level:{tc.chunk_level},cost_time:{tc.chunk_time_length}+\n");
        }
    }

    //这个方法用于将合适的任务分配给建筑工人
    public static Time_Chunk Chunk_Allocator(Building_Unit_Pool bup, Builder_Pipeline_Pool bpp) 
    {
        foreach (Building_Unit bu in bup.bu_set)
        {
            foreach (Builder_Pipeline bp in bpp.bp_set)
            {
                if (bu.avalible_time <= bp.time_length && bu.tc_set.Count != 0)
                {
                    Time_Chunk a = (Time_Chunk)bu.tc_set[0];
                    bp.tc_set.Add(bu.tc_set[0]);
                    Time_Chunk ab = new Time_Chunk() { chunk_id = a.chunk_id, chunk_name = a.chunk_name, chunk_level = a.chunk_level, chunk_time_length = a.chunk_time_length };
                    bp.tc_set.Add(ab);
                    Console.WriteLine($"Builder{bp.pipeline_id} gets the time chunk:{ab.chunk_name}, id:{ab.chunk_id}, level:{ab.chunk_level}");
                    bu.avalible_time += a.chunk_time_length;
                    bp.time_length += a.chunk_time_length;
                    bu.tc_set.RemoveAt(0);
                    return ab;
                }
                else { }
            }
        }
        Time_Chunk c = new Time_Chunk();
        c.chunk_name = "nothing";
        return c;
    }

    //这个方法用于找出任务最少的建筑工人，以将下一个任务分配给他
    public static Builder_Pipeline Find_Bp(Builder_Pipeline_Pool bpp) 
    {
        if (bpp != null)
        {
            bpp.bp_set.Sort();
            Builder_Pipeline bp = bpp.bp_set[0];
            return bp;
        }
        return (Builder_Pipeline)bpp.bp_set[0];
    }      
}

class Builder_Pipeline: IComparable<Builder_Pipeline>
{
    public decimal time_length = 0;
    public int pipeline_id;
    public ArrayList tc_set = new ArrayList();
    public Builder_Pipeline(int id) { pipeline_id = id; }

    public int CompareTo(Builder_Pipeline other)
    {
        if (this.time_length > other.time_length)  return 1; 
        if (this.time_length < other.time_length)  return -1; 
        else  return 0; 
    }
}

class Builder_Pipeline_Pool
{
    public int bp_num;
    public decimal total_time = 0;
    public List<Builder_Pipeline> bp_set = new List<Builder_Pipeline>();
    public Builder_Pipeline_Pool(int number) 
    { 
        bp_num = number; 
        for (int i = 0; i < bp_num; i++) 
        { 
            Builder_Pipeline bp = new Builder_Pipeline(i + 1);
            bp_set.Add(bp);
        }
    }
}

