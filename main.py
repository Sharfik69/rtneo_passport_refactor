from clustering import Clustering
from find_templates import Finder

a = Clustering('templates.xlsx')

a.read_data()

a.clustering()