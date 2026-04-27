# coding: utf-8
"""LF Electrical - Acoplar (Modo Conector Falso)
Insere uma família intermediária para simular uma conexão MEP."""

__title__ = "Acoplar"
__author__ = "Luís Fernando"

from pyrevit import forms, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Electrical import ElectricalSystem
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import clr
import sys, os

lib_path = os.path.dirname(os.path.dirname(__file__)) + "\\LF Electrical.pulldown\\lib"
if lib_path not in sys.path:
    sys.path.append(lib_path)

try:
    import lf_electrical_core
    from lf_electrical_core import doc, uidoc, dbg
except ImportError:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

class AllElementsFilter(ISelectionFilter):
    def AllowElement(self, e): return True
    def AllowReference(self, ref, pos): return False

def get_open_connectors(elem):
    connectors = []
    try:
        mgr = None
        if hasattr(elem, "MEPModel") and elem.MEPModel:
            mgr = elem.MEPModel.ConnectorManager
        elif hasattr(elem, "ConnectorManager"):
            mgr = elem.ConnectorManager
        if mgr:
            for c in mgr.Connectors:
                if not c.IsConnected:
                    connectors.append(c)
    except: pass
    return connectors

def acoplar_com_conector_falso():
    dbg.section("ACOPLAR (Conector Falso) - Início")
    
    try:
        # 1. Selecionar o MESTRE (ex: Quadro, que não tem o conector certo)
        forms.toast("Selecione o elemento MESTRE (ex: Quadro que receberá a conexão)", title="Acoplar")
        ref_master = uidoc.Selection.PickObject(ObjectType.Element, AllElementsFilter(), "Selecione o MESTRE")
        elem_master = doc.GetElement(ref_master.ElementId)
        
        # 2. Selecionar o ALVO (ex: Eletrocalha, Eletroduto)
        forms.toast("Selecione o ALVO (ex: Eletrocalha com ponta solta)", title="Acoplar")
        ref_target = uidoc.Selection.PickObject(ObjectType.Element, AllElementsFilter(), "Selecione o ALVO (Eletrocalha/Eletroduto)")
        elem_target = doc.GetElement(ref_target.ElementId)
        
        target_conns = get_open_connectors(elem_target)
        if not target_conns:
            forms.alert("O elemento ALVO não possui conectores livres para acoplar.", exitscript=True)
            
        # 3. Escolher a família de "Conector Falso"
        # Vamos listar familias de conexoes e equipamentos eletricos
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol).WhereElementIsElementType()
        
        adapter_dict = {}
        for sym in collector:
            if sym.Category:
                c_id = sym.Category.Id.IntegerValue
                if c_id in [int(BuiltInCategory.OST_CableTrayFitting), 
                            int(BuiltInCategory.OST_ConduitFitting),
                            int(BuiltInCategory.OST_ElectricalFixtures),
                            int(BuiltInCategory.OST_GenericModel)]:
                    fname = sym.FamilyName
                    sname = sym.Name
                    adapter_dict["{} - {}".format(fname, sname)] = sym
                    
        if not adapter_dict:
            forms.alert("Nenhuma família adequada encontrada para usar como conector falso.", exitscript=True)
            
        chosen_name = forms.SelectFromList.show(sorted(adapter_dict.keys()), 
                                              title="Selecione a Família do Conector Falso",
                                              button_name="Usar como Acoplador")
        if not chosen_name: return
        
        adapter_sym = adapter_dict[chosen_name]
        
        # Iniciar processo
        with Transaction(doc, "Acoplar com Conector Falso") as t:
            t.Start()
            
            if not adapter_sym.IsActive:
                adapter_sym.Activate()
                doc.Regenerate()
                
            # Local base do mestre
            loc_master = None
            if hasattr(elem_master, "Location") and hasattr(elem_master.Location, "Point"):
                loc_master = elem_master.Location.Point
            elif hasattr(elem_master, "Location") and hasattr(elem_master.Location, "Curve"):
                loc_master = elem_master.Location.Curve.Evaluate(0.5, True)
                
            if not loc_master:
                # Fallback para BoundingBox
                bb = elem_master.get_BoundingBox(doc.ActiveView)
                if bb: loc_master = (bb.Min + bb.Max) / 2.0
                else: loc_master = XYZ.Zero
                
            # Encontrar o conector mais próximo do mestre no alvo
            best_target_conn = target_conns[0]
            min_dist = best_target_conn.Origin.DistanceTo(loc_master)
            for c in target_conns[1:]:
                d = c.Origin.DistanceTo(loc_master)
                if d < min_dist:
                    min_dist = d
                    best_target_conn = c
                    
            # Inserir o conector falso na ponta da eletrocalha/alvo
            insert_pt = best_target_conn.Origin
            
            level_id = elem_target.LevelId if hasattr(elem_target, "LevelId") else elem_master.LevelId
            level = doc.GetElement(level_id) if level_id != ElementId.InvalidElementId else None
            
            fake_conn_inst = None
            if level:
                fake_conn_inst = doc.Create.NewFamilyInstance(insert_pt, adapter_sym, level, Structure.StructuralType.NonStructural)
            else:
                fake_conn_inst = doc.Create.NewFamilyInstance(insert_pt, adapter_sym, Structure.StructuralType.NonStructural)
                
            doc.Regenerate()
            
            # Tentar conectar fisicamente
            fake_conns = get_open_connectors(fake_conn_inst)
            connected = False
            for fc in fake_conns:
                try:
                    best_target_conn.ConnectTo(fc)
                    connected = True
                    break
                except:
                    pass
                    
            if connected:
                dbg.ok("Conector falso fisicamente ligado ao alvo.")
            else:
                dbg.warn("Não foi possível conectar fisicamente o conector falso ao alvo.")
                
            # Agrupar o Mestre e o Conector Falso para se moverem juntos
            try:
                group_ids = clr.Reference[List[ElementId]](List[ElementId]())
                group_ids.Value.Add(elem_master.Id)
                group_ids.Value.Add(fake_conn_inst.Id)
                
                new_group = doc.Create.NewGroup(group_ids.Value)
                dbg.ok("Grupo criado com sucesso.")
            except Exception as e_grp:
                dbg.warn("Não foi possível agrupar os elementos: " + str(e_grp))
                
            t.Commit()
            
        forms.toast("Conexão falsa criada e acoplada ao elemento!", title="Acoplar", warn_icon=False)
        
    except Exception as ex:
        if "OperationCanceledException" in str(type(ex)): pass
        else: forms.alert("Erro:\n" + str(ex))

if __name__ == "__main__":
    acoplar_com_conector_falso()
