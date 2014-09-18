using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ICSharpCode.SharpZipLib.Zip;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace CoordAnalyseService
{
    public class Analyser
    {
        static readonly Regex regex = new Regex("^330([0-9]{3})20([0-9]{6})$");
        public static void ProcessNext()
        {
            try
            {
                var factory = new AccessWorkspaceFactoryClass();
                var ws = (IFeatureWorkspace) factory.OpenFromFile(ConfigurationManager.AppSettings["MdbFile"],0);
                var fc = ws.OpenFeatureClass("Projects");

                using (var conn = new MySqlConnection(ConfigurationManager.ConnectionStrings["LC"].ConnectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText =
                            "SELECT ID,FileName,SavePath FROM uploadfiles WHERE Type = 10 AND State = 0 ORDER BY CreateTime LIMIT 0,1";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read() == false) return;

                            var fileName = reader[1].ToString();
                            var SavePath = reader[2].ToString();
                            var id = Convert.ToInt32(reader[0]);

                            var dest = string.Format("{0}\\Temp\\{1}", ConfigurationManager.AppSettings["BaseFolder"],
                                                     Guid.NewGuid());
                            var msg = string.Empty;
                            try
                            {
                                (new FastZip()).ExtractZip(
                                string.Format("{0}\\{1}", ConfigurationManager.AppSettings["BaseFolder"],
                                              SavePath), dest, "*.txt"
                                );
                                var folder = new DirectoryInfo(dest);
                                ProcessFolder(folder, 5, fc, conn);
                            }
                            catch (Exception ex)
                            {
                                msg = ex.ToString();
                                if (msg.Length > 255) msg = msg.Substring(0, 255);
                            }

                            using (var cmd2 = conn.CreateCommand())
                            {
                                cmd2.CommandText =
                                    string.Format(
                                        "UPDATE uploadfiles SET State = {0}, ProcessMessage = '{1}' WHERE ID={2}",
                                        msg == string.Empty ? "1" : "2", msg, id);
                                cmd2.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var logerror = log4net.LogManager.GetLogger("logerror");
                logerror.ErrorFormat("数据库操作时发生错误：{0}", ex);
            }
            
        }

        private static void ProcessFolder(DirectoryInfo folder, int level, IFeatureClass fc, IDbConnection conn)
        {
            var files = folder.GetFiles("*.txt");
            foreach (var file in files)
            {
                ProcessFile(file, fc, conn);
            }

            if (level > 0)
            {
                var dirs = folder.GetDirectories();
                foreach (var dir in dirs)
                {
                    ProcessFolder(dir, level - 1, fc, conn);
                }
            }
        }

        private static void ProcessFile(FileInfo file, IFeatureClass fc, IDbConnection conn)
        {
            var fileName = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
            
            if (regex.IsMatch(fileName) == false) return;

            var msgs = new List<string>();
            var ret = CheckAll(file.FullName, fc, conn, msgs);

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = string.Format("Update coord_projects SET Note='{0}', Result=1, Visible=1 WHERE ID='{1}'", 
                    string.Join(";", msgs.ToArray()), fileName);
                cmd.ExecuteNonQuery();
            }
        }

        public static bool CheckAll(string txtPath, IFeatureClass fc, IDbConnection conn, IList<string> msgs)
        {
            var list = new List<IPolygon>();
            var projectNo = string.Empty;
            var msg = string.Empty;
            if (LoadPolygon(txtPath, list, out projectNo, out msg) == false)
            {
                msgs.Add(msg);
                return false;
            }

            if(CheckArea(list, projectNo, conn, out msg) == false) msgs.Add(msg);

            try
            {
                if(CheckOverlay(list, projectNo, fc, out msg) == false) msgs.Add(msg);
                UpdatePolygons(list, projectNo, fc);
            }
            catch (Exception ex)
            {
                msgs.Add("检查过程中出现错误：" + ex);
            }

            return msgs.Count == 0;
        }

        private static bool LoadPolygon(string txtPath, IList<IPolygon> polygons, out string projectNo,
                                       out string msg)
        {
            projectNo = System.IO.Path.GetFileNameWithoutExtension(txtPath);
            var missing = Type.Missing;
            var pgCount = 0;
            try
            {
                using (var reader = new StreamReader(txtPath, Encoding.GetEncoding("GB2312")))
                {
                    var row = 0;
                    var begin = false;
                    var line = reader.ReadLine();
                    row++;
                    if (line == null) line = string.Empty;
                    line = line.Trim();
                    var lastTokens = new[] {"", "", "", "", "", "", "", "", ""};

                    var polygon = new PolygonClass();

                    var ringId = string.Empty;
                    var ring = new RingClass();
                    var pc = (IPointCollection) ring;

                    while (line != null)
                    {
                        line = line.Trim();
                        if (line == "[地块坐标]")
                        {
                            begin = true;
                            line = reader.ReadLine();
                            row++;
                            continue;
                        }

                        if (begin && string.IsNullOrEmpty(line) == false)
                        {
                            string[] tokens = line.Split(',');
                            if (tokens[tokens.Length - 1] == "@")
                            {
                                if (tokens.Length != 9)
                                {
                                    msg = string.Format("文件第{0}行的格式有误，字段数量应为9。", row);
                                    return false;
                                }

                                if (pc.PointCount > 0)
                                {
                                    if (pc.PointCount < 4)
                                    {
                                        msg = string.Format("文件第{0}行之前的坐标串有误，多边形顶点数不应少于4个", row);
                                        pc = (IPointCollection) ring;
                                        return false;
                                    }
                                    else
                                    {
                                        polygon.AddGeometry(ring, ref missing, ref missing);
                                    }
                                }

                                if (polygon.GeometryCount > 0)
                                {
                                    polygons.Add(polygon);
                                    pgCount++;
                                    polygon = new PolygonClass();
                                    ring = new RingClass();
                                    pc = (IPointCollection) ring;
                                    ringId = string.Empty;

                                }
                                lastTokens = tokens;
                            }
                            else
                            {
                                if (tokens.Length == 4)
                                {
                                    IPoint pt = new PointClass();
                                    double x, y;
                                    if (tokens[1] != ringId)
                                    {
                                        if (pc.PointCount > 0)
                                        {
                                            polygon.AddGeometry(ring, ref missing, ref missing);
                                        }
                                        ring = new RingClass();
                                        pc = (IPointCollection) ring;
                                        ringId = tokens[1];
                                    }


                                    if ((!double.TryParse(tokens[2], out y) || !double.TryParse(tokens[3], out x)) ||
                                        x < 40000000 || x > 41000000 || y < 2000000 || y > 4000000)
                                    {
                                        msg = string.Format("文件第{0}行的格式有误，坐标应为数字类型。", row);
                                        polygon = new PolygonClass();
                                        pc = (IPointCollection) polygon;
                                        return false;
                                    }

                                    pt.Y = y;
                                    pt.X = x;
                                    pc.AddPoint(pt, ref missing, ref missing);
                                }
                                else
                                {
                                    msg = string.Format("文件第{0}行的格式有误，字段数量应为4。", row);
                                    polygon = new PolygonClass();
                                    ring = new RingClass();
                                    pc = (IPointCollection) ring;
                                    return false;
                                }
                            }
                        }

                        line = reader.ReadLine();
                        row++;

                    }


                    if (pc.PointCount > 0)
                    {
                        if (pc.PointCount < 4)
                        {
                            msg = string.Format("文件第{0}行之前的坐标串有误，多边形顶点数不应少于4个", row);
                            return false;
                        }
                        else
                        {
                            polygon.AddGeometry(ring, ref missing, ref missing);
                        }
                    }


                    if (polygon.GeometryCount > 0)
                    {
                        try
                        {
                            polygons.Add(polygon);

                            pgCount++;
                        }
                        catch (Exception ex)
                        {
                            msg = string.Format("导入第{0}个地块时发生错误: {1}", pgCount, ex.Message);
                            return false;
                        }
                    }


                    if (pgCount > 0)
                    {
                        msg = string.Format("成功导入{0}个图斑", pgCount);
                        return true;
                    }
                    else
                    {
                        msg = string.Format("文件中无图斑");
                        return false;
                    }

                }
            }
            catch (Exception ex)
            {
                msg =  string.Format("导入第{0}个地块时发生错误: {1}", pgCount, ex.Message);
                return false;
            }
        }

        private static void UpdatePolygons(IList<IPolygon> polygons, string projectNo, IFeatureClass fc)
        {
            var table = (ITable) fc;
            table.DeleteSearchedRows(new QueryFilterClass()
                {
                    WhereClause = string.Format("ProjectNo='{0}'", projectNo)
                });

            var cursor = fc.Insert(true);
            var index = cursor.FindField("ProjectNo");
            foreach (var pg in polygons)
            {
                var buffer = fc.CreateFeatureBuffer();
                buffer.Shape = pg;
                buffer.set_Value(index, projectNo);
                cursor.InsertFeature(buffer);
            }
            Marshal.ReleaseComObject(cursor);
        }

        private static bool CheckOverlay(IList<IPolygon> polygons, string projectNo, IFeatureClass fc, out string msg)
        {
            foreach (var pg in polygons)
            {
                (pg as IPolygon4).SimplifyEx(true, true, true);
                var to = (ITopologicalOperator)pg;
                var filter = new SpatialFilterClass()
                    {
                        Geometry = pg,
                        WhereClause = string.Format("ProjectNo <> '{0}'", projectNo),
                        SpatialRel = esriSpatialRelEnum.esriSpatialRelEnvelopeIntersects
                    };

                var cursor = fc.Search(filter, true);
                var feature = cursor.NextFeature();
                while (feature != null)
                {
                    var shape = feature.ShapeCopy;
                    if (shape != null || shape.IsEmpty == false ||
                        shape.GeometryType == esriGeometryType.esriGeometryPolygon)
                    {
                        var intersect = to.Intersect(shape, esriGeometryDimension.esriGeometry2Dimension);
                        if (intersect is IArea && (intersect as IArea).Area > double.Epsilon)
                        {
                            msg = "与其他项目图斑相交";
                            return false;
                        }
                    }

                    feature = cursor.NextFeature();
                }
                Marshal.ReleaseComObject(cursor);
            }
            msg = string.Empty;
            return true;
        }

        private static bool CheckArea(IList<IPolygon> polygons, string projectNo, IDbConnection conn, out string msg)
        {
            var area = 0.0;
            foreach (var pg in polygons)
            {
                area = Math.Abs((pg as IArea).Area);
            }

            area /= 10000;

            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = string.Format("Select Area from projects where ID = '{0}'", projectNo);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.IsDBNull(0))
                        {
                            msg = string.Format("已上传的自查表中找不到项目'{0}'，请先上传自查表", projectNo);
                            return false;
                        }
                        else
                        {
                            var area2 = Convert.ToDouble(reader[0]);
                            if (Math.Abs((area - area2)/area2) >= 0.2)
                            {
                                msg = "图斑面积与项目总规模相差超过20%";
                                return false;
                            }
                            else
                            {
                                msg = string.Empty;
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                msg = string.Format("在已上传的自查表中查询项目面积失败:" + ex);
                return false;
                
            }
            
        }
    }

}
