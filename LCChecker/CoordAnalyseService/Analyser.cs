using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.esriSystem;
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
using System.Threading;

namespace CoordAnalyseService
{
    public class Analyser
    {
        static readonly Regex regex = new Regex("^33([0-9]{4})([0-9]{8})$");
        public static void ProcessNext()
        {
            try
            {
                //var factory = new AccessWorkspaceFactoryClass();
                //var ws = (IFeatureWorkspace)factory.OpenFromFile(AppDomain.CurrentDomain.BaseDirectory + "\\" + ConfigurationManager.AppSettings["MDBFile"], 0);
                var factory = new FileGDBWorkspaceFactoryClass();
                var ws = (IFeatureWorkspace)factory.OpenFromFile(AppDomain.CurrentDomain.BaseDirectory + "\\" + ConfigurationManager.AppSettings["MDBFile"], 0);
                var fc = ws.OpenFeatureClass("Projects");
                var to = ws.OpenFeatureClass("CulProjects");

                using (var conn = new MySqlConnection(ConfigurationManager.ConnectionStrings["LC"].ConnectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = string.Format(
                            "SELECT ID,FileName,SavePath,CityID,Type FROM uploadfiles WHERE (Type = 10 or Type = 11) AND State = 0 {0} ORDER BY CreateTime LIMIT 0,1",
                            string.IsNullOrEmpty(ConfigurationManager.AppSettings["AdditionalCondition"]) ? string.Empty : (" AND " + ConfigurationManager.AppSettings["AdditionalCondition"]));
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read() == false) return;


                            var fileName = reader[1].ToString();
                            var savePath = reader[2].ToString();
                            var id = Convert.ToInt32(reader[0]);
                            var cityId = Convert.ToInt32(reader[3]);
                            var type = Convert.ToInt32(reader[4]);
                           
                            reader.Close();
                            var shortPath = System.IO.Path.GetFileName(savePath);
                            Console.WriteLine(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][{1}]开始:{2}", DateTime.Now, cityId, shortPath));
                            string msg;
                            if (type == 10)
                            {
                                ProcessProjectRange(fileName, savePath, id, cityId, fc, conn, out msg);
                            }
                            else
                            {
                                ProcessCultivateRange(fileName, savePath, id, cityId, fc, to, conn, out msg);
                            }
                                
                            using (var cmd2 = conn.CreateCommand())
                            {
                                cmd2.CommandText =
                                    string.Format(
                                        "UPDATE uploadfiles SET State = {0}, ProcessMessage = '{1}' WHERE ID={2}",
                                        msg == string.Empty ? "1" : "2", msg, id);
                                cmd2.ExecuteNonQuery();
                            }
                            if (msg == string.Empty)
                                Console.WriteLine(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][{1}]成功", DateTime.Now, cityId));
                            else
                                Console.WriteLine(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][{1}]失败:{2}", DateTime.Now, cityId, msg));
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

        #region 新增耕地范围线
        private static void ProcessCultivateRange(string filename, string savePath, int id, int cityId, IFeatureClass fc, IFeatureClass to, IDbConnection conn, out string msg)
        {
            var dest = string.Format("{0}\\Spatial\\{1}", ConfigurationManager.AppSettings["BaseFolder"],
                                                    Guid.NewGuid());
            msg = string.Empty;
            try
            {
                var src = string.Format("{0}\\{1}", ConfigurationManager.AppSettings["BaseFolder"],
                                         savePath);

                (new FastZip()).ExtractZip(src
                , dest, null
                );
                if (Directory.Exists(dest))
                {
                    var folder = new DirectoryInfo(dest);
                    ProcessFolder2(folder, 5, cityId, fc, to, conn);
                }
                else
                {
                    msg = "压缩包内容为空或者不可识别。";
                }
            }
            catch (ZipException)
            {
                msg = "压缩包内容不可识别，请更换压缩软件";
            }
            catch (ApplicationException ex)
            {
                msg = ex.Message;
            }
            catch (Exception ex)
            {
                msg = ex.ToString();
                if (msg.Length > 255) msg = msg.Substring(0, 255);
            }
        }

        private static bool CheckCity2(string projectNo, int cityId, IDbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    string.Format(
                        "SELECT count(1) FROM coord_newareaprojects WHERE ID='{0}' AND CityID = {1}",
                        projectNo, cityId);
                return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            }
        }

        private static void ProcessFolder2(DirectoryInfo folder, int level, int cityId, IFeatureClass fc, IFeatureClass to, IDbConnection conn)
        {
            var files = folder.GetFiles("*.shp");
            foreach (var file in files)
            {
                ProcessFile2(file, cityId, fc, to,  conn);
            }

            if (level > 0)
            {
                var dirs = folder.GetDirectories();
                foreach (var dir in dirs)
                {
                    ProcessFolder2(dir, level - 1, cityId, fc,to, conn);
                }
            }
        }

        private static void ProcessFile2(FileInfo file, int cityId, IFeatureClass fc, IFeatureClass to, IDbConnection conn)
        {
            var fileName = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
            CheckAll2(file.FullName, fc, to, conn, cityId);
        }

        public static void CheckAll2(string shpPath, IFeatureClass fc, IFeatureClass to, IDbConnection conn, int cityId)
        {
            var cfc = CheckShapefile(shpPath, to);
            var projects = GetAllProjects(cfc, cityId, conn);

            foreach (var project in projects)
            {
                CheckProject(project, cfc, fc, to, conn);
            }
        }

        public static void CheckProject(string projectNo, IFeatureClass cfc, IFeatureClass fc, IFeatureClass to, IDbConnection conn)
        {
            var msg = string.Empty;
            var ret = false;
            if (CheckProjectRangeDone(projectNo, conn))
            {
                var ret1 = CheckProperties(projectNo, cfc, conn, out msg);
                var ret2 = CheckRange(projectNo, cfc, fc);
                if (ret1 && ret2)
                {
                    ret = true;
                }
                else
                {
                    if (ret2 == false) msg += "新增耕地坐标必须在项目范围之内。";
                }
            }
            else
            {
                msg = "项目范围坐标存疑，无法进一步检查新增耕地坐标。";
            }

            if (ret)
            {
                CopyRecords(projectNo, cfc, to);
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = string.Format("Update coord_newareaprojects SET Note='{0}', Result={2}, Visible=1, UpdateTime=now() WHERE ID='{1}'",
                    MySQLEncode(msg), MySQLEncode(projectNo), ret ? "1" : "0");
                cmd.ExecuteNonQuery();
            }
        }
        
        /// <summary>
        /// 判断Shapefile格式是否正确
        /// </summary>
        /// <param name="shpPath"></param>
        /// <returns></returns>
        public static IFeatureClass CheckShapefile(string shpPath, IFeatureClass to)
        {
            try
            {
                var folder = System.IO.Path.GetDirectoryName(shpPath);
                var factory = new ShapefileWorkspaceFactoryClass();
                var ws = (IFeatureWorkspace)factory.OpenFromFile(folder, 0);
                
                var fc = ws.OpenFeatureClass(System.IO.Path.GetFileNameWithoutExtension(shpPath));
                var ds = (IGeoDataset)fc;
                var ds2 = (IGeoDataset)to;
                if (((IClone)(ds.SpatialReference)).IsEqual((IClone)(ds2.SpatialReference)) == false)
                {
                    throw new ApplicationException("文件的坐标系统错误，请确保是西安80坐标高斯投影3度分带带号40，东偏4050000(Xian_1980_3_Degree_GK_Zone_40)");
                }

                var fields = new[] { "DIKUAIAREA", "DIKUAI_NO", "MAP_NO", "PL_NAME", "XMBH", "GDZLDB", "DLBM" };
                foreach (var fld in fields)
                {
                    if (fc.FindField(fld) < 0)
                    {
                        throw new ApplicationException("文件中缺少字段: " + fld);
                    }
                }

                var intFields = new[] { "DIKUAI_NO", "GDZLDB" };
                foreach (var fld in intFields)
                {
                    var field = fc.Fields.Field[fc.FindField(fld)];
                    if (field.Type != esriFieldType.esriFieldTypeInteger && field.Type != esriFieldType.esriFieldTypeSmallInteger)
                    {
                        throw new ApplicationException("文件中" + fld + "字段不是整形");
                    }
                }

                var strFields = new[] { "MAP_NO", "PL_NAME", "XMBH", "DLBM" };
                foreach (var fld in strFields)
                {
                    var field = fc.Fields.Field[fc.FindField(fld)];
                    if (field.Type != esriFieldType.esriFieldTypeString)
                    {
                        throw new ApplicationException("文件中" + fld + "字段不是文本类型");
                    }
                }

                var floatFields = new[] { "DIKUAIAREA" };
                foreach (var fld in floatFields)
                {
                    var field = fc.Fields.Field[fc.FindField(fld)];
                    if (field.Type != esriFieldType.esriFieldTypeDouble && field.Type != esriFieldType.esriFieldTypeSingle)
                    {
                        throw new ApplicationException("文件中" + fld + "字段不是浮点类型");
                    }
                }
                return fc;
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            catch (Exception)
            {
                throw new ApplicationException("打开文件错误: " + System.IO.Path.GetFileName(shpPath));
            }
        }

        /// <summary>
        /// 拷贝信息
        /// </summary>
        /// <param name="projectNo"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        public static void CopyRecords(string projectNo, IFeatureClass from, IFeatureClass to)
        {
            var table = (ITable)to;
           
            table.DeleteSearchedRows(new QueryFilterClass()
            {
                WhereClause = string.Format("XMBH='{0}'", projectNo)
            });

            var cursor2 = to.Insert(true);
            var cursor = from.Search(new QueryFilterClass { WhereClause = string.Format("XMBH='{0}'", projectNo) }, true);
            var feature = cursor.NextFeature();
            var fields = new[] { "DIKUAIAREA", "DIKUAI_NO", "MAP_NO", "PL_NAME", "XMBH", "GDZLDB", "DLBM" };
            var list = new List<int>();
            var list2 = new List<int>();
            foreach (var fld in fields) list.Add(cursor.FindField(fld));
            foreach (var fld in fields) list2.Add(cursor2.FindField(fld));
            
            while (feature != null)
            {
                var buffer = to.CreateFeatureBuffer();
                buffer.Shape = feature.ShapeCopy;
                for(var i = 0;i<list.Count;i++)
                {
                    var index = list[i];
                    var index2 = list2[i];
                
                    buffer.set_Value(index2, feature.get_Value(index));
                }
                cursor2.InsertFeature(buffer);
                feature = cursor.NextFeature();
            }

            Marshal.ReleaseComObject(cursor);
            Marshal.ReleaseComObject(cursor2);
        }

        public static IList<string> GetAllProjects(IFeatureClass cfc, int cityId, IDbConnection conn)
        {
            var dict = new Dictionary<string, string>();
            var cursor2 = cfc.Search(null, true);
            var feature = cursor2.NextFeature();
            
            var index = cfc.FindField("XMBH");
            while (feature != null)
            {
                var projectNo = feature.get_Value(index).ToString();
                if (regex.IsMatch(projectNo) == false)
                {
                    throw new ApplicationException("shp文件中的项目编号(XMBH)错误:" + projectNo);
                }

                if (dict.ContainsKey(projectNo) == false)
                {
                    if (CheckCity2(projectNo, cityId, conn))
                    {
                        dict.Add(projectNo, projectNo);
                    }
                }

                feature = cursor2.NextFeature();
            }

            Marshal.ReleaseComObject(cursor2);
            return dict.Values.ToList();
        }

        public static bool CheckProjectRangeDone(string projectNo, IDbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT count(1) FROM coord_projects WHERE ID = '" + projectNo + "' AND (Visible = 0 OR Exception = 1 OR Result = 1)";
                var count = Convert.ToInt32(cmd.ExecuteScalar());
                return count > 0;
                
            }
        }
        
        /// <summary>
        /// 检查耕地的属性，包含总面积、水田旱地面积、等别等
        /// </summary>
        /// <param name="projectNo"></param>
        /// <param name="cfc"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static bool CheckProperties(string projectNo, IFeatureClass cfc, IDbConnection conn, out string msg)
        {
            var threshold = double.Parse(ConfigurationManager.AppSettings["NewAreaTolerance"]);


            var level1 = 0;
            var area1 = 0.0;

            var level2 = 0;
            var area2 = 0.0;
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "SELECT Type,Degree,Area FROM farmland WHERE ProjectID = '" + projectNo + "'";
                using (var reader = cmd.ExecuteReader())
                {
                    while(reader.Read())
                    {
                        if (Convert.ToInt32(reader[0]) == 2)
                        {
                            level2 = Convert.ToInt32(reader[1]);
                            area2 = Math.Round(Convert.ToDouble(reader[2]),4);
                        }
                        else
                        {
                            level1 = Convert.ToInt32(reader[1]);
                            area1 = Math.Round(Convert.ToDouble(reader[2]),4);
                        }
                    }
                    
                }
            }

            var alevel1 = 0;
            var aarea1 = 0.0;
            var count1 = 0;

            var alevel2 = 0;
            var aarea2 = 0.0;
            var count2 = 0;

            var cursor2 = cfc.Search(new QueryFilterClass { WhereClause = string.Format("XMBH='{0}'", projectNo) }, true);
            var feature = cursor2.NextFeature();
            
            var index1 = cfc.FindField("GDZLDB");
            var index2 = cfc.FindField("DLBM");
            while (feature != null)
            {
                var geo = feature.Shape;
                
                if (geo.IsEmpty == false && geo is IArea)
                {
                    var area = (geo as IArea).Area;
                    var type = feature.get_Value(index2).ToString();
                    if (type == "水田" || type == "011")
                    {
                        aarea1 += area;
                        alevel1 += Convert.ToInt32(feature.get_Value(index1));
                        count1++;
                    }
                    else if (type == "旱地" || type == "013")
                    {
                        aarea2 += area;
                        alevel2 += Convert.ToInt32(feature.get_Value(index1));
                        count2++;
                    }
                }
                feature = cursor2.NextFeature();
            }

            Marshal.ReleaseComObject(cursor2);
            aarea1 = Math.Round(aarea1/10000, 4);
            aarea2 = Math.Round(aarea2/10000, 4);
            if (area1 > aarea1 + 0.0001)
            {
                msg = string.Format("水田图斑面积({0}公顷)必须大于等于填报面积({1}公顷)", aarea1, area1);
                return false;
            }
            else if ((count1 == 0 && level1 > 0) || (count1 > 0 && level1 == 0))
            {
                msg = "图斑中水田情况与表9填报情况不符";
                return false;
            }
            else if (area1 > double.Epsilon && Math.Abs(area1 - aarea1) / area1 > threshold)
            {
                msg = string.Format("水田图斑面积({0}公顷)比填报面积({1}公顷)大10%以上", aarea1, area1);
                return false;
            }
            else if (count1 != 0 && Math.Abs(alevel1 * 1.0 / count1 - level1 * 1.0) >= 1.0)
            {
                msg = string.Format("水田图斑的耕地质量等别(等别{0})与表9中填表情况(等别{1})不符", alevel1 / count1, level1);
                return false;
            }

            if (area2 > aarea2 + 0.0001)
            {
                msg = string.Format("旱地图斑面积({0}公顷)必须大于等于填报面积({1}公顷)", aarea2, area2);
                return false;
            }
            else if((count2 == 0  && level2 > 0) ||(count2>0&&level2 == 0))
            {
                msg = "图斑中旱地情况与表9填报情况不符";
                return false;
            }
            else if (area2>double.Epsilon && Math.Abs(area2 - aarea2) / area2 > threshold)
            {
                msg = string.Format("旱地图斑面积({0}公顷)比填报面积({1}公顷)大10%以上", aarea2, area2);
                return false;
            }
            else if (count2 != 0 && Math.Abs( alevel2 * 1.0 / count2 - level2 * 1.0)>=1.0)
            {
                msg = string.Format("旱地图斑的耕地质量等别(等别{0})与表9中填表情况(等别{1})不符", alevel2/count2, level2);
                return false;
            }
            msg = string.Empty;
            return true;
        }

        /// <summary>
        /// 检查新增耕地是否在项目范围内
        /// </summary>
        /// <param name="projectNo"></param>
        /// <param name="cfc"></param>
        /// <param name="fc"></param>
        public static bool CheckRange(string projectNo, IFeatureClass cfc, IFeatureClass fc)
        {
            var threshold = double.Parse(ConfigurationManager.AppSettings["NewOverlayTolerance"]);
            var ranges = new List<IPolygon>();
            var cursor = fc.Search(new QueryFilterClass { WhereClause = string.Format("ProjectNo='{0}'", projectNo) }, true);
            var feature = cursor.NextFeature();
            while (feature != null)
            {
                var geo = feature.ShapeCopy;

                if (geo.IsEmpty == false && geo is IPolygon)
                {
                    var pg = (IPolygon)geo;
                    (pg as IPolygon4).SimplifyEx(true, true, true);
                    if (pg.IsEmpty == false)
                    {
                        ranges.Add(pg);
                    }
                }

                feature = cursor.NextFeature();
            }

            Marshal.ReleaseComObject(cursor);

            var cursor2 = cfc.Search(new QueryFilterClass { WhereClause = string.Format("XMBH='{0}'", projectNo) }, true);
            feature = cursor2.NextFeature();
            var ret = true;
            while (feature != null)
            {
                var geo = feature.Shape;

                if (geo.IsEmpty == false && geo is IPolygon)
                {
                    var pg = (IPolygon)geo;
                    if (InRange(pg, ranges, threshold) == false)
                    {
                        ret = false;
                        break;
                    }
                }

                feature = cursor2.NextFeature();
            }

            Marshal.ReleaseComObject(cursor2);
            return ret;
        }

        private static bool InRange(IPolygon pg, List<IPolygon> ranges, double threshold)
        {
            var polygon = pg as IPolygon4;
            IGeometryBag exteriorRings = polygon.ExteriorRingBag;

            IEnumGeometry exteriorRingsEnum = exteriorRings as IEnumGeometry;
            exteriorRingsEnum.Reset();
            IRing currentExteriorRing = exteriorRingsEnum.Next() as IRing;
            
            while (currentExteriorRing != null)
            {
                if (InRange(currentExteriorRing, polygon, ranges, threshold) == false) return false;
                currentExteriorRing = exteriorRingsEnum.Next() as IRing;
            }

            return true;
        }

        private static bool InRange(IRing ring, IPolygon4 polygon, List<IPolygon> ranges, double threshold)
        {
            var pg = new PolygonClass();
            var missing = Type.Missing;
            pg.AddGeometry(ring, ref missing, ref missing);
            var bag = polygon.get_InteriorRingBag(ring);
            var interiorEnum = bag as IEnumGeometry;
            interiorEnum.Reset();
            var r = interiorEnum.Next() as IRing;

            while (r != null)
            {
                pg.AddGeometry(r, ref missing, ref missing);
                r = interiorEnum.Next() as IRing;
            }

            pg.ITopologicalOperator2_IsKnownSimple_2 = false;
            (pg as ITopologicalOperator2).Simplify();

            foreach (var range in ranges)
            {
                var to = (ITopologicalOperator)range;
                var intersect = to.Intersect(pg, esriGeometryDimension.esriGeometry2Dimension);
                if (intersect is IArea && pg is IArea && Math.Abs((intersect as IArea).Area - (pg as IArea).Area) < threshold)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion

        #region 项目范围线检查

        private static void ProcessProjectRange(string filename, string savePath, int id, int cityId, IFeatureClass fc, IDbConnection conn, out string msg)
        {
            var dest = string.Format("{0}\\Spatial\\{1}", ConfigurationManager.AppSettings["BaseFolder"],
                                                    Guid.NewGuid());
            msg = string.Empty;
            try
            {
                var src = string.Format("{0}\\{1}", ConfigurationManager.AppSettings["BaseFolder"],
                                         savePath);

                (new FastZip()).ExtractZip(src
                , dest, ".txt$"
                );
                if (Directory.Exists(dest))
                {
                    var folder = new DirectoryInfo(dest);
                    var projectNoDict = new Dictionary<string, string>();
                    msg = CheckFolder(folder, 5, cityId, projectNoDict, conn);

                    if (string.IsNullOrEmpty(msg))
                    {
                        ProcessFolder(folder, 5, cityId, projectNoDict, fc, conn);
                    }
                }
                else
                {
                    msg = "压缩包内容为空或者不可识别。";
                }
            }
            catch (Exception ex)
            {
                msg = ex.ToString();
                if (msg.Length > 255) msg = msg.Substring(0, 255);
            }
        }

        private static bool CheckCity(string projectNo, int cityId, IDbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    string.Format(
                        "SELECT count(1) FROM coord_projects WHERE ID='{0}'" ,// AND CityID = {1}",
                        projectNo, cityId);
                return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            }
        }

        private static string CheckFolder(DirectoryInfo folder, int level, int cityId, Dictionary<string,string> projectDict, IDbConnection conn)
        {
            var files = folder.GetFiles("*.txt");
            foreach (var file in files)
            {
                var filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
                if (regex.IsMatch(filename) == false)
                {
                    return "文件名必须为项目备案编号：" + filename;
                }
                else if(CheckCity(filename, cityId, conn) == false)
                {
                    return "项目列表中不存在当前项目：" + filename + ",请联系维护人员";
                }
                else if (IsInSecondBatch(filename, conn) == false)
                {
                    return "不是第二批项目：" + filename;
                }
                else
                {
                    projectDict[filename] = filename;
                }
            }

            if (level > 0)
            {
                var dirs = folder.GetDirectories();
                foreach (var dir in dirs)
                {
                    var msg = CheckFolder(dir, level - 1, cityId, projectDict, conn);
                    if (string.IsNullOrEmpty(msg) == false) return msg;
                }
            }
            else
            {
                return "压缩包中目录结构太深，不能多于5层。";
            }

            return string.Empty;
        }

        private static bool IsInSecondBatch(string projectNo, IDbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = string.Format("SELECT count(1) FROM seprojects WHERE ID='{0}'", projectNo);
                var count = Convert.ToInt32(cmd.ExecuteScalar());
                return count > 0;
            }
        }

        private static void ProcessFolder(DirectoryInfo folder, int level, int cityId, Dictionary<string, string> projectNoDict, IFeatureClass fc, IDbConnection conn)
        {
            var files = folder.GetFiles("*.txt");
            foreach (var file in files)
            {
                ProcessFile(file, cityId,projectNoDict, fc, conn);
            }

            if (level > 0)
            {
                var dirs = folder.GetDirectories();
                foreach (var dir in dirs)
                {
                    ProcessFolder(dir, level - 1, cityId, projectNoDict, fc, conn);
                }
            }
        }

        private static string MySQLEncode(string str)
        {
            str = str.Replace("'", @"\'");
            str = str.Replace("\"", "\\\"");
            return str;
        }

        private static void ProcessFile(FileInfo file, int cityId, Dictionary<string, string> projectNoDict, IFeatureClass fc, IDbConnection conn)
        {
            var fileName = System.IO.Path.GetFileNameWithoutExtension(file.FullName);

            if (projectNoDict.ContainsKey(fileName) == false) return;

            /*if (regex.IsMatch(fileName) == false) return;

            if (isInFirstBatch(fileName, fc) == true) return;*/

            var msgs = new List<string>();
            var ret = CheckAll(file.FullName, projectNoDict, fc, conn, msgs);
            if(ret) TryCopyFile(file, cityId);
            using (var cmd = conn.CreateCommand())
            {
                if (IsException(fileName, conn))
                {
                    cmd.CommandText = string.Format("Update coord_projects SET Note='{0}', Visible=1, UpdateTime=now() WHERE ID='{1}' and CityID={2}", 
                    MySQLEncode(string.Join(";", msgs.ToArray())), MySQLEncode(fileName), cityId);
                }
                else
                {
                    cmd.CommandText = string.Format("Update coord_projects SET Note='{0}', Result={2}, Visible=1, UpdateTime=now() WHERE ID='{1}' and CityID={3}",
                    MySQLEncode(string.Join(";", msgs.ToArray())), MySQLEncode(fileName), ret ? "1" : "0", cityId);
                }
                
                cmd.ExecuteNonQuery();
            }
        }

        private static void TryCopyFile(FileInfo file, int cityId)
        {
            var baseFolder = ConfigurationManager.AppSettings["BaseFolder"];
            for (var i = 0; i < 3; i++)
            {
                try
                {
                    file.CopyTo(string.Format(@"{0}\App_Data\CoordProject\{1}\{2}", baseFolder, cityId, file.Name), true);
                    return;
                }
                catch (IOException ex)
                {
                    Console.WriteLine(string.Format("拷贝文件'{0}'失败：{1}", file.Name, ex));
                    Thread.Sleep(2000);
                }
            }
        }

        /*
        public static bool isInFirstBatch(string txtPath, IFeatureClass fc)
        {
            var projectNo = System.IO.Path.GetFileNameWithoutExtension(txtPath);
            var cursor = fc.Search(new QueryFilterClass { WhereClause = string.Format("ProjectNo='{0}'", projectNo) }, true);

            var f = cursor.NextFeature();
            var ret = false;
            if(f != null)
            {
                ret = f.get_Value(f.Fields.FindField("Batch")).ToString() == "1";
            }
            Marshal.ReleaseComObject(cursor);
            return ret;
        }*/

        public static bool IsException(string projectNo, IDbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = string.Format("SELECT Exception FROM coord_projects WHERE ID='{0}'", projectNo);
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read() == false) return false;

                    return Convert.ToBoolean(reader[0]);
                }
            
            }
        }

        public static bool CheckAll(string txtPath, Dictionary<string, string> projectNoDict, IFeatureClass fc, IDbConnection conn, IList<string> msgs)
        {
            var list = new List<IPolygon>();
            var projectNo = string.Empty;
            var msg = string.Empty;
            if (LoadPolygon(txtPath, list, out projectNo, out msg) == false)
            {
                if (projectNoDict.ContainsKey(projectNo)) projectNoDict.Remove(projectNo);
                msgs.Add(msg);
                return false;
            }

            if (IsException(projectNo, conn) == false)
            {
                if (CheckArea(list, projectNo, conn, out msg) == false) msgs.Add(msg);
                try
                {
                    if (IsHighLevel(projectNo, conn) == false && CheckOverlay(list, projectNo, projectNoDict, fc, conn, out msg) == false)  msgs.Add(msg);
                }
                catch (Exception ex)
                {
                    msgs.Add("检查过程中出现错误：" + ex);
                }
            }

            if (msgs.Count == 0)
            {
                try
                {
                    UpdatePolygons(list, projectNo, fc, conn);
                }
                catch (Exception ex)
                {
                    msgs.Add("更新库时发生错误：" + ex);
                }
            }

            projectNoDict.Remove(projectNo);

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
                                if (tokens.Length <7)
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
                                    polygon.ITopologicalOperator2_IsKnownSimple_2 = false;
                                    (polygon as ITopologicalOperator2).Simplify();
                                    polygon.GeometriesChanged();
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
                            polygon.ITopologicalOperator2_IsKnownSimple_2 = false;
                            (polygon as ITopologicalOperator2).Simplify();
                            polygon.GeometriesChanged();
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
                        msg = string.Format("文件中无图斑，或未严格遵照部格式。");
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

        private static void UpdatePolygons(IList<IPolygon> polygons, string projectNo, IFeatureClass fc, IDbConnection conn)
        {
            var table = (ITable) fc;
            table.DeleteSearchedRows(new QueryFilterClass()
                {
                    WhereClause = string.Format("ProjectNo='{0}'", projectNo)
                });

            var cursor = fc.Insert(true);
            var index = cursor.FindField("ProjectNo");
            var index2 = cursor.FindField("HighLevel");
            var index3 = cursor.FindField("Batch");
            foreach (var pg in polygons)
            {
                var buffer = fc.CreateFeatureBuffer();
                buffer.Shape = pg;
                buffer.set_Value(index, projectNo);
                buffer.set_Value(index2, IsHighLevel(projectNo, conn) ? 1 : 0);
                buffer.set_Value(index3, IsInSecondBatch(projectNo, conn) ? 2 : 1);
                cursor.InsertFeature(buffer);
            }
            Marshal.ReleaseComObject(cursor);
        }

        private static bool IsHighLevel(string projectNo, IDbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT Name FROM coord_projects WHERE ID='" + projectNo + "'";
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read() == false) return false;

                    var name = Convert.ToString(reader[0]);
                    return name.Contains("高标准");
                }
            }

            /*var filter = new QueryFilterClass()
                    {
                        
                        WhereClause = string.Format("ProjectNo ='{0}'", projectNo),
                    };

            var cursor = fc.Search(filter, true);

            var index = fc.FindField("HighLevel");
            IFeature f = cursor.NextFeature();
            if(f == null)return false;

            var ret = (f.get_Value(index) == DBNull.Value) ? false : Convert.ToInt32(f.get_Value(index)) == 1;
            Marshal.ReleaseComObject(cursor);
            return ret;*/    
        }

        private static bool CheckOverlay(IList<IPolygon> polygons, string projectNo, Dictionary<string, string> projectNoDict, IFeatureClass fc, IDbConnection conn, out string msg)
        {
            var index = fc.FindField("ProjectNo");
            var index2 = fc.FindField("Batch");
            
            //var list = new List<string>();
            var dict = new Dictionary<string, double>();
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
                    var projNo = feature.get_Value(index).ToString();

                    if (IsHighLevel(projNo, conn) == false)
                    {
                        var firstBatch = !IsInSecondBatch(projNo, conn);// feature.get_Value(index2).ToString() == "1";
                        if (firstBatch) projNo += "(第一批)";

                        if (projectNoDict.ContainsKey(projNo) == false)
                        {
                            var shape = feature.ShapeCopy;

                            if (shape != null || shape.IsEmpty == false ||
                                shape.GeometryType == esriGeometryType.esriGeometryPolygon)
                            {
                                var intersect = to.Intersect(shape, esriGeometryDimension.esriGeometry2Dimension);
                                if (intersect is IArea && (intersect as IArea).Area > double.Epsilon)
                                {
                                    var area = (intersect as IArea).Area;
                                    if (dict.ContainsKey(projNo))
                                    {
                                        dict[projNo] = dict[projNo] + area;
                                    }
                                    else
                                    {
                                        dict.Add(projNo, area);
                                    }
                                    //area += (intersect as IArea).Area;
                                    //if (list.Exists(x => x == projNo) == false) list.Add(projNo);
                                }
                            }
                        }
                    }
                
                    feature = cursor.NextFeature();
                }
                Marshal.ReleaseComObject(cursor);
            }

            var threshold = double.Parse(ConfigurationManager.AppSettings["OverlayTolerance"]);
            var list = dict.Where(x=>x.Value>=threshold).ToList();

            //if (area < double.Parse(ConfigurationManager.AppSettings["OverlayTolerance"]))
            if(list.Count == 0)
            {
                msg = string.Empty;
                return true;
            }
            else
            {
                if (list.Count > 5) list.RemoveRange(5, list.Count - 5);
                msg = "与其他项目图斑相交，项目编号："+ string.Join(",", list.Select(x=>x.Key).ToArray());
                return false;
            }
        }

        private static bool CheckArea(IList<IPolygon> polygons, string projectNo, IDbConnection conn, out string msg)
        {
            var area = 0.0;
            foreach (var pg in polygons)
            {
                area = area + Math.Round(Math.Abs((pg as IArea).Area),4);
            }

            area = Math.Round(area / 10000, 4);
            var tol = double.Parse(ConfigurationManager.AppSettings["AreaTolerance"]);
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = string.Format("Select Area from seprojects where ID = '{0}'", projectNo);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            if (reader.IsDBNull(0))
                            {
                                msg = string.Format("已上传的自查表中找不到项目'{0}'，请先上传自查表", projectNo);
                                return false;
                            }
                            else
                            {
                                var area2 = Convert.ToDouble(reader[0]);
                                if (Math.Abs((area - area2)/area2) >= tol)
                                {
                                    msg = string.Format("图斑面积({1}公顷)与项目总规模({2}公顷)相差超过{0}%", tol * 100, area, Math.Round(area2, 4));
                                    return false;
                                }
                                else
                                {
                                    msg = string.Empty;
                                    return true;
                                }
                            }
                        }
                        else
                        {
                            msg = string.Format("已上传的自查表中找不到项目'{0}'，请先上传自查表", projectNo);
                            return false;
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
        #endregion
    }

}
