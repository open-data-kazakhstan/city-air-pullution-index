from datapackage import Package

package = Package()

package.infer(r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll_piv.csv')
package.infer(r'C:\Users\USER\Desktop\practice\city-air-pullution-index\data\air_poll.csv')

package.commit()
package.save(r'C:\Users\USER\Desktop\practice\city-air-pullution-index\datapackage.json')
